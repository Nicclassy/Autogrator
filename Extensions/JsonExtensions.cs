using Newtonsoft.Json.Linq;

namespace Autogrator.Extensions;

public static class JsonExtensions {
    public static void Walk(this JToken token, Action<JProperty> action) {
        if (token.Type == JTokenType.Object) {
            foreach (JProperty child in token.Children<JProperty>()) {
                action(child);
                Walk(child.Value, action);
            }
        } else if (token.Type == JTokenType.Array) {
            foreach (JToken child in token.Children()) {
                Walk(child, action);
            }
        }
    }
}