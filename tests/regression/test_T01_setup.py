from kinea_pptx_plus.io import extract_chart_xml


def test_can_load_chart_xml():
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    assert b"<c:chartSpace" in xml
    assert b"lineChart" in xml.lower() or b"linechart" in xml.lower()
