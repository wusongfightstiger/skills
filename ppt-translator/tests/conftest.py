import pytest
from pathlib import Path
import csv
from datetime import date


class MockConfig:
    """Mock configuration for testing without API key."""

    minimax_api_host = "api.minimax.chat"
    minimax_model = "MiniMax-Text-01"
    request_timeout = 120
    default_font = "Arial"
    font_size_min = 8
    font_size_shrink_ratio = 0.5

    def __init__(self):
        """Initialize mock configuration."""
        self.api_key = "mock_api_key_for_testing"
        self.glossary_path = Path("/mock/glossary.csv")


@pytest.fixture
def mock_config():
    """Provide a mock configuration for testing."""
    return MockConfig()


@pytest.fixture
def test_glossary_csv(tmp_path):
    """Create a temporary test glossary CSV."""
    csv_path = tmp_path / "test_glossary.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["中文术语", "英文翻译", "领域", "添加日期", "是否已确认"])
        writer.writeheader()
        writer.writerows([
            {"中文术语": "电阻", "英文翻译": "resistor", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
            {"中文术语": "电容", "英文翻译": "capacitor", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
            {"中文术语": "欧拉定律", "英文翻译": "Euler's Law", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
        ])
    return csv_path


@pytest.fixture
def test_pptx(tmp_path):
    """Create a minimal test PPTX file."""
    pptx_path = tmp_path / "test.pptx"

    slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="1" name="TextBox 1"/>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="1000000" cy="500000"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst/>
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="zh-CN" dirty="0"/>
                            <a:t>电阻</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>'''

    import zipfile
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
        zf.writestr("[Content_Types].xml", content_types)

    return pptx_path
