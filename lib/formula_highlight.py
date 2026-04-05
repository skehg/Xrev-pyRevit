# -*- coding: utf-8 -*-
"""
FormulaEditorHighlightMixin
---------------------------
Mixin for WPF windows that contain a RichTextBox named ``txtFormula`` and
expose ``self._autocomplete_names`` (a list of str).

Provides:
  - Text-access wrappers replacing the plain TextBox API
      _formula_get_text()
      _formula_set_text(text)
      _formula_get_caret()
      _formula_set_caret(idx)
      _formula_get_selection_start()
      _formula_get_selection_length()
      _formula_select(start, length)
  - Syntax-highlight engine
      _apply_syntax_highlights()

Usage
-----
1. Inherit from this mixin alongside forms.WPFWindow.
2. Set  self._highlighting = False  in __init__ BEFORE calling
   _apply_syntax_highlights() for the first time.
3. Call _apply_syntax_highlights() inside on_formula_changed and
   on_formula_selection_changed event handlers.
"""

import re
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Thickness
from System.Windows.Documents import (
    LogicalDirection,
    Paragraph,
    Run,
    TextPointerContext,
    TextRange,
)
from System.Windows.Media import Color, SolidColorBrush

# TextPointerContext.None is a reserved word in Python; access via getattr.
_TPC_NONE = getattr(TextPointerContext, "None")


# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------

def _parse_hex_color(hex_str):
    """Parse '#RRGGBB' or '#RGB' into a System.Windows.Media.Color."""
    h = hex_str.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return Color.FromRgb(r, g, b)


def _lighter_color(hex_str, factor=0.65):
    """Lerp a hex colour toward white by *factor* (0 = unchanged, 1 = white)."""
    h = hex_str.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r2 = r + int((255 - r) * factor)
    g2 = g + int((255 - g) * factor)
    b2 = b + int((255 - b) * factor)
    return Color.FromRgb(r2, g2, b2)


def _solid(color):
    return SolidColorBrush(color)


# ---------------------------------------------------------------------------
# Mixin
# ---------------------------------------------------------------------------

class FormulaEditorHighlightMixin(object):
    """
    Add-in to a WPFWindow subclass that provides coloured syntax highlighting
    inside a RichTextBox named ``txtFormula``.

    Contract
    --------
    The host class MUST provide:
      self.txtFormula          - the WPF RichTextBox control
      self._autocomplete_names - list[str] of family parameter names
      self._highlighting       - bool flag (initialise to False in __init__)
    """

    # ------------------------------------------------------------------
    # Generic RichTextBox text-access helpers (control passed as argument)
    # ------------------------------------------------------------------

    def _rtb_get_text(self, rtb):
        """Return the plain-text content of *rtb*, without trailing newlines."""
        doc = rtb.Document
        full = TextRange(doc.ContentStart, doc.ContentEnd)
        return full.Text.rstrip("\r\n")

    def _rtb_set_text(self, rtb, text):
        """Replace the entire content of *rtb* with *text*."""
        doc = rtb.Document
        full = TextRange(doc.ContentStart, doc.ContentEnd)
        full.Text = text if text is not None else ""

    def _rtb_get_caret(self, rtb):
        """Walk text nodes only to count plain-text chars up to CaretPosition."""
        doc = rtb.Document
        caret = rtb.CaretPosition
        if caret is None:
            return 0
        count = 0
        ptr = doc.ContentStart
        while ptr is not None:
            if ptr.CompareTo(caret) >= 0:
                break
            ctx = ptr.GetPointerContext(LogicalDirection.Forward)
            if ctx == _TPC_NONE:
                break
            if ctx == TextPointerContext.Text:
                next_ptr = ptr.GetNextContextPosition(LogicalDirection.Forward)
                # caret lands within this text run
                if next_ptr is None or caret.CompareTo(next_ptr) < 0:
                    count += len(TextRange(ptr, caret).Text)
                    break
                count += ptr.GetTextRunLength(LogicalDirection.Forward)
                ptr = next_ptr
            else:
                ptr = ptr.GetNextContextPosition(LogicalDirection.Forward)
        return count

    def _rtb_set_caret(self, rtb, idx):
        """Move the caret of *rtb* to char offset *idx*."""
        ptr = self._ptr_at_char_for(rtb, idx)
        if ptr is not None:
            rtb.Selection.Select(ptr, ptr)

    def _rtb_get_selection_start(self, rtb):
        """Return the selection start of *rtb* as a plain-text char offset."""
        doc = rtb.Document
        sel_start = rtb.Selection.Start
        if sel_start is None:
            return 0
        return len(TextRange(doc.ContentStart, sel_start).Text)

    def _rtb_get_selection_length(self, rtb):
        """Return the selection length of *rtb* in plain-text chars."""
        sel = rtb.Selection
        if sel is None:
            return 0
        return len(sel.Text)

    def _rtb_select(self, rtb, start, length):
        """Select *length* chars starting at offset *start* in *rtb*."""
        ptr_s = self._ptr_at_char_for(rtb, start)
        ptr_e = self._ptr_at_char_for(rtb, start + length)
        if ptr_s is not None and ptr_e is not None:
            rtb.Selection.Select(ptr_s, ptr_e)

    def _ptr_at_char_for(self, rtb, idx):
        """
        Return a TextPointer at plain-text char offset *idx* from
        *rtb*.ContentStart, or ContentEnd if *idx* is out of range.
        """
        if idx < 0:
            idx = 0
        doc = rtb.Document
        ptr = doc.ContentStart
        remaining = idx

        while ptr is not None:
            ctx = ptr.GetPointerContext(LogicalDirection.Forward)

            if ctx == _TPC_NONE:
                break

            if ctx == TextPointerContext.Text:
                run_len = ptr.GetTextRunLength(LogicalDirection.Forward)
                if remaining <= run_len:
                    return ptr.GetPositionAtOffset(remaining)
                remaining -= run_len
                ptr = ptr.GetNextContextPosition(LogicalDirection.Forward)

            elif ctx in (
                TextPointerContext.ElementStart,
                TextPointerContext.ElementEnd,
                TextPointerContext.EmbeddedElement,
            ):
                ptr = ptr.GetNextContextPosition(LogicalDirection.Forward)

            else:
                break

        # idx beyond end — return ContentEnd
        return doc.ContentEnd

    # ------------------------------------------------------------------
    # txtFormula convenience delegates (keep existing call-sites simple)
    # ------------------------------------------------------------------

    def _formula_get_text(self):
        return self._rtb_get_text(self.txtFormula)

    def _formula_set_text(self, text):
        self._rtb_set_text(self.txtFormula, text)

    def _formula_get_caret(self):
        return self._rtb_get_caret(self.txtFormula)

    def _formula_set_caret(self, idx):
        self._rtb_set_caret(self.txtFormula, idx)

    def _formula_get_selection_start(self):
        return self._rtb_get_selection_start(self.txtFormula)

    def _formula_get_selection_length(self):
        return self._rtb_get_selection_length(self.txtFormula)

    def _formula_select(self, start, length):
        self._rtb_select(self.txtFormula, start, length)

    def _ptr_at_char(self, idx):
        return self._ptr_at_char_for(self.txtFormula, idx)

    # ------------------------------------------------------------------
    # Syntax highlighting engine
    # ------------------------------------------------------------------

    def _bracket_color_hex(self, ch):
        """Return an '#RRGGBB' fg colour for a bracket character."""
        if ch in "()":
            return "#0060C0"
        if ch in "[]":
            return "#1A7A1A"
        if ch in "{}":
            return "#B06000"
        return "#444444"

    def _tokenize_formula(self, text, caret_idx):
        """
        Return list of (segment_str, fg_Color_or_None, bg_Color_or_None).
        Lower priority: parameter name tokens → blue foreground.
        Higher priority: the bracket char adjacent to the caret and its
        matching counterpart → coloured fg + light bg.
        """
        n = len(text)
        if n == 0:
            return []

        # per-character colour arrays (None = default)
        fg_map = [None] * n
        bg_map = [None] * n

        # --- pass 0: mark characters inside double-quoted strings (no highlighting) ---
        quoted_map = [False] * n
        in_quote = False
        for i, ch in enumerate(text):
            if ch == '"':
                quoted_map[i] = True
                in_quote = not in_quote
            elif in_quote:
                quoted_map[i] = True

        # --- pass 1: parameter name tokens (case-sensitive word boundaries) ---
        names = sorted(self._autocomplete_names, key=lambda s: -len(s))
        for name in names:
            if not name:
                continue
            try:
                pattern = r"(?<![A-Za-z0-9_]){}(?![A-Za-z0-9_])".format(re.escape(name))
                for m in re.finditer(pattern, text):
                    col = _parse_hex_color("#0000C8")
                    for i in range(m.start(), m.end()):
                        if fg_map[i] is None and not quoted_map[i]:
                            fg_map[i] = col
            except Exception:
                pass

        # --- pass 2: matching bracket pair adjacent to caret (higher priority) ---
        brackets_open = "([{"
        brackets_close = ")]}"
        all_brackets = brackets_open + brackets_close

        # check the character just before the caret and the one at the caret
        candidate_indices = []
        if 0 < caret_idx <= n:
            candidate_indices.append(caret_idx - 1)
        if 0 <= caret_idx < n:
            candidate_indices.append(caret_idx)

        for ci in candidate_indices:
            if 0 <= ci < n and text[ci] in all_brackets and not quoted_map[ci]:
                match_idx = self._find_matching_bracket(text, ci)
                if match_idx is not None:
                    # Use a strong highlight so bracket pairing is clearly visible.
                    fg_col = _parse_hex_color("#111111")
                    bg_col = _parse_hex_color("#FFF59D")
                    for bi in (ci, match_idx):
                        if 0 <= bi < n:
                            fg_map[bi] = fg_col
                            bg_map[bi] = bg_col
                    break  # only process the innermost adjacent bracket

        # --- build segments: merge consecutive chars with identical colour ---
        segments = []
        i = 0
        while i < n:
            fg = fg_map[i]
            bg = bg_map[i]
            j = i + 1
            while j < n and fg_map[j] == fg and bg_map[j] == bg:
                j += 1
            segments.append((text[i:j], fg, bg))
            i = j

        return segments

    def _apply_syntax_highlights_to(self, rtb):
        """
        Rebuild *rtb* document with syntax-coloured Run objects.
        Called from formula-changed/selection-changed handlers for any
        RichTextBox formula editor in the host window.
        """
        if getattr(self, "_highlighting", False):
            return

        self._highlighting = True
        try:
            caret_idx = self._rtb_get_caret(rtb)
            sel_start = self._rtb_get_selection_start(rtb)
            sel_length = self._rtb_get_selection_length(rtb)
            text = self._rtb_get_text(rtb)

            doc = rtb.Document
            doc.Blocks.Clear()

            para = Paragraph()
            para.Margin = Thickness(0)

            if text:
                segments = self._tokenize_formula(text, caret_idx)
                for seg_text, fg, bg in segments:
                    run = Run(seg_text)
                    if fg is not None:
                        run.Foreground = _solid(fg)
                    if bg is not None:
                        run.Background = _solid(bg)
                    para.Inlines.Add(run)

            doc.Blocks.Add(para)

            # Restore selection if one was active, otherwise just restore caret.
            if sel_length > 0:
                self._rtb_select(rtb, sel_start, sel_length)
            else:
                self._rtb_set_caret(rtb, caret_idx)

        finally:
            self._highlighting = False

    def _apply_syntax_highlights(self):
        """Convenience wrapper — highlights self.txtFormula."""
        self._apply_syntax_highlights_to(self.txtFormula)
