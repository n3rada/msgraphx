# tests/test_html.py

from __future__ import annotations

from rich.markdown import Markdown

from msgraphx.utils.html import html_to_markdown, render_html, strip_html


def test_mention_renders_bold():
    assert html_to_markdown('<p>hi <at id="0">Bob</at></p>') == "hi **@Bob**"


def test_attachment_renders_placeholder():
    md = html_to_markdown('<p>see<attachment id="abc"></attachment>below</p>')
    assert "[attachment]" in md


def test_img_renders_alt_text():
    assert "[diagram]" in html_to_markdown('<p><img src="x.png" alt="diagram"></p>')


def test_img_without_alt_uses_default():
    assert "[image]" in html_to_markdown('<p><img src="x.png"></p>')


def test_spans_are_stripped():
    assert html_to_markdown("<p><span>plain</span> text</p>") == "plain text"


def test_render_html_returns_markdown():
    assert isinstance(render_html("<p>hi <at>Bob</at></p>"), Markdown)


def test_strip_html_collapses_whitespace_and_entities():
    assert strip_html("<div>a  &amp;<br>  b</div>") == "a & b"
