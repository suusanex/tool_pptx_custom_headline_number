using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxHeadlineNumbering;

public sealed record ParagraphInfo(
    int SlideIndex,
    string ShapeName,
    P.PlaceholderValues? PlaceholderType,
    int ParagraphLevel,
    string CurrentText,
    A.Paragraph ParagraphElement);
