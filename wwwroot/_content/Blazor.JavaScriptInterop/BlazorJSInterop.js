// This file is to show how a library package may provide JavaScript interop features
// wrapped in a .NET API

window.BlazorJSFunctions =
{
    ShowPrompt: function (message)
    {
        return prompt(message, 'Type anything here');
    },   
    CopyText: function (text)
    {
        // original value
        var returnValue = 0;

        try
        {
            navigator.clipboard.writeText(text);

            // set to 1;
            returnValue = 1;
        }
        catch (err)
        {
            returnValue = -2;
        }  

        // return value
        return returnValue;
    }
};
