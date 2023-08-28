using Microsoft.JSInterop;
using System.Drawing;

namespace Blazor.Excelerate
{
    public class BlazorJSBridge
    {
        public static ValueTask<string> Prompt(IJSRuntime jsRuntime, string message)
        {
            // Implemented in BlazorJSInterop.js
            return jsRuntime.InvokeAsync<string>(
                "BlazorJSFunctions.ShowPrompt",
                message);
        }

        public async static Task<int> CopyToClipboard(IJSRuntime jsRuntime, string textToCopy)
        {
            // set the value
            int copied = await jsRuntime.InvokeAsync<int>("BlazorJSFunctions.CopyText", textToCopy);

            // return value
            return copied;
        }
    }
}
