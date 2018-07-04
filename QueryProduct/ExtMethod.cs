using System;
using System.Windows.Forms;

namespace QueryProduct
{
    public static class ExtMethod
    {
        public static object Invoke(this Control control, Action mDelegate)
        {
            return control.Invoke(mDelegate);
        }
    }
}