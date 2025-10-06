using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Контекст обработки SolidWorks документа
/// </summary>
public class SolidWorksDocumentContext
{
    public ModelDoc2 Model { get; set; } = null!;
    public DrawingDoc? Drawing { get; set; }
    public SldWorks Application { get; set; } = null!;
}