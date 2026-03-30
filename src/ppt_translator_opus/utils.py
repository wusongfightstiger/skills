"""Parallel translation scheduler with retry, rate-limit handling, and progress monitoring."""

import asyncio
from datetime import datetime
import time
import sys

import httpx

from .engines.base import TranslationEngine


class ProgressMonitor:
    """Track translation progress and report every N seconds."""

    def __init__(self, total: int, interval: int = 600):
        self.total = total
        self.interval = interval  # default 10 minutes
        self.success = 0
        self.failed = 0
        self.start_time = time.time()
        self._task: asyncio.Task | None = None

    @property
    def remaining(self) -> int:
        return self.total - self.success - self.failed

    @property
    def completed(self) -> int:
        return self.success + self.failed

    def record_success(self):
        self.success += 1

    def record_failure(self):
        self.failed += 1

    def _estimate_remaining_time(self) -> str:
        elapsed = time.time() - self.start_time
        if self.completed == 0:
            return "估算中..."
        avg_per_item = elapsed / self.completed
        eta_seconds = avg_per_item * self.remaining
        if eta_seconds < 60:
            return f"{eta_seconds:.0f} 秒"
        elif eta_seconds < 3600:
            return f"{eta_seconds / 60:.0f} 分钟"
        else:
            hours = int(eta_seconds // 3600)
            mins = int((eta_seconds % 3600) // 60)
            return f"{hours} 小时 {mins} 分钟"

    def report(self):
        now = datetime.now().strftime("%H:%M")
        elapsed = time.time() - self.start_time
        elapsed_str = f"{elapsed / 60:.1f} 分钟" if elapsed >= 60 else f"{elapsed:.0f} 秒"
        eta = self._estimate_remaining_time()

        print(
            f"\n{'='*50}\n"
            f"  进度报告 [{now}]  已运行 {elapsed_str}\n"
            f"{'='*50}\n"
            f"  ✓ 成功: {self.success}\n"
            f"  ✗ 失败: {self.failed}\n"
            f"  … 剩余: {self.remaining}\n"
            f"  预计还需: {eta}\n"
            f"{'='*50}\n",
            flush=True,
        )

    async def _monitor_loop(self):
        """Background task that reports progress at fixed intervals."""
        while self.remaining > 0:
            await asyncio.sleep(self.interval)
            if self.remaining > 0:
                self.report()

    def start(self):
        self._task = asyncio.create_task(self._monitor_loop())

    def stop(self):
        if self._task and not self._task.done():
            self._task.cancel()


async def translate_all_slides(
    slides: list[dict],
    engine: TranslationEngine,
    glossary: list[dict],
    max_concurrent: int = 5,
    progress_interval: int = 600,
) -> list[dict | None]:
    """Translate all slides in parallel using the given engine.

    Args:
        slides: List of slide dicts to translate.
        engine: Translation engine to use.
        glossary: Glossary terms.
        max_concurrent: Max parallel API calls.
        progress_interval: Seconds between progress reports (default 600 = 10 min).

    Returns a list of translated slide dicts (or None for failed slides).
    """
    semaphore = asyncio.Semaphore(max_concurrent)
    consecutive_429 = 0
    rate_limit_delay = 0.0  # shared delay applied before each API call
    total = len(slides)
    monitor = ProgressMonitor(total, interval=progress_interval)

    async def translate_one(slide: dict, idx: int) -> dict | None:
        nonlocal consecutive_429, rate_limit_delay

        async with semaphore:
            for attempt in range(4):  # initial + 3 retries
                # Apply shared rate-limit delay
                if rate_limit_delay > 0:
                    await asyncio.sleep(rate_limit_delay)
                try:
                    result = await engine.translate_slide(slide, glossary)
                    consecutive_429 = 0
                    monitor.record_success()
                    print(
                        f"  [{monitor.success + monitor.failed}/{total}] "
                        f"Slide {slide['slide_number']} translated "
                        f"({len(slide['elements'])} elements)",
                        flush=True,
                    )
                    return result
                except httpx.HTTPStatusError as e:
                    if e.response.status_code == 429:
                        consecutive_429 += 1
                        # Progressive slowdown: increase shared delay on consecutive 429s
                        if consecutive_429 >= 3:
                            rate_limit_delay = min(rate_limit_delay + 5.0, 30.0)
                            consecutive_429 = 0
                            print(
                                f"  Rate limited, adding {rate_limit_delay:.0f}s delay between requests",
                                file=sys.stderr,
                                flush=True,
                            )
                        # Exponential backoff for this specific retry
                        retry_after = e.response.headers.get("Retry-After")
                        if retry_after:
                            wait = float(retry_after)
                        else:
                            wait = 2 ** attempt  # 1, 2, 4, 8
                        if attempt < 3:
                            print(
                                f"  Slide {slide['slide_number']}: 429, retrying in {wait:.0f}s (attempt {attempt + 1}/3)",
                                file=sys.stderr,
                                flush=True,
                            )
                            await asyncio.sleep(wait)
                        continue
                    else:
                        print(
                            f"  Slide {slide['slide_number']}: HTTP {e.response.status_code}",
                            file=sys.stderr,
                            flush=True,
                        )
                        monitor.record_failure()
                        return None
                except Exception as e:
                    print(
                        f"  Slide {slide['slide_number']}: {type(e).__name__}: {e}",
                        file=sys.stderr,
                        flush=True,
                    )
                    monitor.record_failure()
                    return None

            # All retries exhausted
            print(
                f"  Slide {slide['slide_number']}: failed after 3 retries",
                file=sys.stderr,
                flush=True,
            )
            monitor.record_failure()
            return None

    # Start progress monitor
    monitor.start()

    try:
        tasks = [translate_one(slide, i) for i, slide in enumerate(slides)]
        results = await asyncio.gather(*tasks)
    finally:
        monitor.stop()

    return list(results)


def merge_translations(originals: list[dict], translated: list[dict | None]) -> list[dict]:
    """Merge translated slides with originals, falling back to original on failure."""
    merged = []
    for orig, trans in zip(originals, translated):
        if trans is not None:
            merged.append(trans)
        else:
            merged.append(orig)
    return merged
