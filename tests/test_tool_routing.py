from types import SimpleNamespace

from backend import agent


def _names(group: str) -> set[str]:
    return {tool["function"]["name"] for tool in agent.tools_for(group)}


def test_explicit_groups_cover_the_registered_tools_once():
    word_tools = set(agent.TOOL_GROUPS["word_all"])
    excel_tools = set(agent.TOOL_GROUPS["excel_all"])

    assert word_tools.isdisjoint(excel_tools)
    assert word_tools | excel_tools == set(agent.ALL_TOOL_NAMES)
    assert _names("merge_word") == {"word_merge_documents"}
    assert _names("merge_excel") == {"merge_files"}
    assert _names("full") == set(agent.ALL_TOOL_NAMES)


def test_routes_word_documents_to_word_tools():
    assert agent.route("edit the second paragraph", ["report.DOCX"]) == ("word_all", False)
    assert "word_create_table" in _names("word_all")
    assert "word_delete_table" in _names("word_all")
    assert "remove_duplicates" not in _names("word_all")


def test_routes_excel_and_csv_to_excel_tools():
    assert agent.route("remove duplicates", ["sales.XLSX", "extra.csv"]) == ("excel_all", False)
    assert "remove_duplicates" in _names("excel_all")
    assert "add_row" in _names("excel_all")
    assert "merge_files" in _names("excel_all")
    assert "word_read_document" not in _names("excel_all")


def test_routes_compatible_file_merges_to_only_the_merge_tool():
    assert agent.route("merge all files", ["a.xlsx", "b.csv"]) == ("merge_excel", False)
    assert agent.route("combine these Word documents", ["a.docx", "b.docx"]) == ("merge_word", False)


def test_mixed_or_unknown_file_types_fall_back_to_full_tool_set():
    assert agent.route("merge these files", ["a.docx", "b.xlsx"]) == ("full", False)
    assert agent.route("sort this sheet", ["a.xlsx", "legacy.xls"]) == ("full", False)
    assert agent.route("make a change", []) == ("create", False)


def test_routes_session_file_objects_and_does_not_misroute_row_append():
    files = [SimpleNamespace(display_name="report.docx"), SimpleNamespace(display_name="table.docx")]
    assert agent.route("append a new paragraph", files) == ("word_all", False)
    assert agent.route("append these documents", files) == ("merge_word", False)


def test_in_file_merge_requests_keep_the_full_format_tool_group():
    assert agent.route("merge cells A1:B2", ["a.xlsx", "b.xlsx"]) == ("excel_all", False)
    assert agent.route("merge the tables in both reports", ["a.docx", "b.docx"]) == ("word_all", False)
