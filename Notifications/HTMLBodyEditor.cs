using AngleSharp;
using AngleSharp.Dom;

namespace Autogrator.Notifications;

public sealed class HTMLBodyEditor {
    private readonly IDocument document;

    public HTMLBodyEditor(string body) {
        IConfiguration config = Configuration.Default.WithDefaultLoader();
        IBrowsingContext context = new BrowsingContext(config);
        this.document = context.OpenAsync(request => request.Content(body)).Result;
    }

    public void AddText(string text) {
        IElement body = document.QuerySelector("body")!;
        IElement div = body.QuerySelector("div")!;
        IElement firstElement = div.FirstElementChild;

        IElement parent = document.CreateElement("p");
        parent.ClassName = "MsoNormal";
        firstElement.InsertBefore(parent);

        IElement textElement = document.CreateElement("o:p");
        textElement.TextContent = text;
        parent.AppendChild(textElement);
    }

    public string Content() => document.DocumentElement.OuterHtml;
}