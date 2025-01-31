using AngleSharp;
using AngleSharp.Dom;

namespace Autogrator.Notifications;

public readonly struct HTMLBodyEditor {
    private readonly IDocument document;

    public HTMLBodyEditor(string body) {
        IConfiguration config = Configuration.Default.WithDefaultLoader();
        IBrowsingContext context = new BrowsingContext(config);
        document = context.OpenAsync(request => request.Content(body)).GetAwaiter().GetResult();
    }

    public void PrependText(string text) {
        IElement body = document.QuerySelector("body")!;
        IElement div = body.QuerySelector("div")!;
        IElement firstElement = div.FirstElementChild!;

        IElement parent = document.CreateElement("p");
        parent.ClassName = "MsoNormal";
        firstElement.InsertBefore(parent);

        IElement textElement = document.CreateElement("o:p");
        textElement.InnerHtml = text.TrimEnd().Replace("\n", "<br>");
        parent.AppendChild(textElement);
    }

    public string Content() => document.DocumentElement.OuterHtml;
}