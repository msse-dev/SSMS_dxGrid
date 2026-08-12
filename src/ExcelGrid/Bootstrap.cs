using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace ExcelGrid.Ssms;

[Export(typeof(IWpfTextViewCreationListener))]
[ContentType("text")]
[TextViewRole(PredefinedTextViewRoles.Document)]
internal sealed class Bootstrap : IWpfTextViewCreationListener
{
    public void TextViewCreated(IWpfTextView textView) => GridDiscovery.Start();
}
