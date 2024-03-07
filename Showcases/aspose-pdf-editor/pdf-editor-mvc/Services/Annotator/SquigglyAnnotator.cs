using Aspose.Pdf;
using aspose.pdf.annotation.Model;
using Aspose.Pdf.Annotations;

namespace aspose.pdf.annotation;

public class SquigglyAnnotator : IAnnotator
{
    private readonly SquigglyModel _model;
    private readonly string _workFolder;
    private readonly string _inputFile;
    private readonly string _outputFile;

    public SquigglyAnnotator(SquigglyModel model, string workFolder, string inputFile, string outputFile)
    {
        _model = model;
        _workFolder = workFolder;
        _inputFile = inputFile;
        _outputFile = outputFile;
    }

    public void Save()
    {
        // Load the PDF file
        Document document = new Document(Path.Combine(_workFolder, _inputFile));

        var textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber(_model.Text);
        textFragmentAbsorber.Visit(document.Pages[_model.PageNumber]);

        SquigglyAnnotation squigglyAnnotation = new SquigglyAnnotation(
            document.Pages[_model.PageNumber],
            textFragmentAbsorber.TextFragments[1].Rectangle)
        {
            Title = _model.Title,
            Subject = _model.Subject,
            Color = _model.ColorValue,
            Opacity = _model.Opacity
        };

        // Add annotation to the page
        document
            .Pages[_model.PageNumber]
            .Annotations
            .Add(squigglyAnnotation);

        document.Save(Path.Combine(_workFolder, _outputFile));
    }
}