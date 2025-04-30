<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/141575675/21.1.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830556)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
# Rich Text Editor for WPF -- How to Use Syntax Highlight Tokens to implement T-SQL language Syntax Highlight

This example illustrates how to register the [ISyntaxHighlightService](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.Services.ISyntaxHighlightService) to implement simplified syntax highlighting for the T-SQL language.

## Implementation Details

Text is parsed into tokens (a list of [SyntaxHighlightToken](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.SyntaxHighlightToken) class instances) manually by the `CustomSyntaxHighlightService.ParseTokens()` method call. The resulting list is sorted and passed to the [SubDocument.ApplySyntaxHighlight](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.SubDocument.ApplySyntaxHighlight(System.Collections.Generic.List-DevExpress.XtraRichEdit.API.Native.SyntaxHighlightToken-)) method.

## More Examples

* [Use DevExpress CodeParser and Syntax Highlight tokens to Highlight C# and VB Code](https://github.com/DevExpress-Examples/syntax-highlighting-for-c-and-vb-code-using-devexpress-codeparser-and-syntax-highlight-tokens)

## Documentation

* [How to: Highlight Document Syntax](https://docs.devexpress.com/WPF/14714/controls-and-libraries/rich-text-editor/examples/automation/how-to-highlight-document-syntax)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-implement-t-sql-language-syntax-highlighting-by-creating-syntax-highlight-tokens&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-implement-t-sql-language-syntax-highlighting-by-creating-syntax-highlight-tokens&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
