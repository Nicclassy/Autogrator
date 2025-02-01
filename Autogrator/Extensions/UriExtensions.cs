using System;

namespace Autogrator.Extensions;

public static class UriExtensions {
    public static Uri Append(this Uri uri, string path) {
        UriBuilder builder = new(uri);
        builder.Path += path;
        return builder.Uri;
    }
}