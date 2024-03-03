using Aspose.Pdf;
using System.Text.Json.Serialization;

namespace aspose.pdf.annotation.Model.Descriptions;

[Serializable]
public class TitleModel1
{
    public string Title { get; set; } = "title";

    public string Subject { get; set; } = "subject";

    [JsonIgnore]
    public Color? ColorValue 
    { 
        get 
        {
            return Aspose.Pdf.Color.Parse(Color);
        } 
    }

    public string Color { get; set; } = "Aqua";

    public double Opacity { get; set; } = 0.5;
}