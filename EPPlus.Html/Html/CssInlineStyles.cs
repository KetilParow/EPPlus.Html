using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class CssInlineStyles : CssDeclaration, IRenderElement
    {
        public void Render(StringBuilder html)
        {
            html.Append("style=\"");
            html.Append(ToString());
            html.Append('\"');
        }
        public override string ToString()
        {
            return string.Join("", this
                    .Where(x => x.Value != null)
                    .Select(x => x.Key + ":" + x.Value + ";"));
        }
    }
}
