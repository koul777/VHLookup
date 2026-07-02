from __future__ import annotations

from openpyxl.comments import Comment


def make_comment(text: str, author: str = "VHLookup") -> Comment:
    comment = Comment(text, author)
    line_count = max(1, str(text).count("\n") + 1)
    comment.width = 420
    comment.height = max(160, min(320, 80 + line_count * 28))
    return comment
