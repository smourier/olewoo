/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using olewoo.interop;

namespace olewoo
{
    public class OWTypeLib : ITlibNode, IClearUp
    {
        ITypeLib _tlib;
        readonly string _name;

        public OWTypeLib(string path)
        {
            NativeMethods.LoadTypeLib(path, out _tlib);
            if (_tlib == null)
                throw new ArgumentException(null, path + " is not a loadable typelibrary.");

            _name = _tlib.GetName();
            _name += " (" + _tlib.GetHelpDocumentation(out int cnt) + ")";
        }

        public void ClearUp()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_tlib);
            _tlib = null;
        }

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => throw new NotSupportedException();
        public override int ImageIndex => (int)ImageIndices.idx_typelib;
        public override string ShortName => _tlib.GetName();
        public override string ObjectName => ShortName;
        public override string Name => _name;
        public override ITlibNode Parent => null;

        public override List<ITlibNode> GenChildren()
        {
            var res = new List<ITlibNode>();
            int ticount = _tlib.GetTypeInfoCount();
            for (int x = 0; x < ticount; ++x)
            {
                _tlib.GetTypeInfo(x, out ITypeInfo ti);
                CommonBuildTlibNode(this, ti, true, true, res);
            }
            return res;
        }

        public override void BuildIDLInto(IDLFormatter ih)
        {
            // Header for type library, followed by a first pass to pre-declare
            // interfaces.
            // dispinterfaces aren't shown seperately.

            ih.AppendLine("// Generated .IDL file (by OleWoo)");
            ih.AppendLine("[");
            var liba = new List<string>();
            var tla = new TypeLibAttr(_tlib);
            liba.Add("uuid(" + tla.guid + ")");
            liba.Add("version(" + tla.wMajorVerNum + "." + tla.wMinorVerNum + ")");

            var cds = new CustomDatas(_tlib as ITypeLib2);
            {
                foreach (var cd in cds.Items)
                {
                    liba.Add("custom(" + cd.guid + ", " + ITypeInfoXtra.QuoteString(cd.varValue) + ")");
                }
            }

            string help = _tlib.GetHelpDocumentation(out int cnt);
            if (!string.IsNullOrEmpty(help))
            {
                liba.Add("helpstring(\"" + help + "\")");
            }

            if (cnt != 0)
            {
                liba.Add("helpcontext(" + cnt.PaddedHex() + ")");
            }

            cnt = 0;
            liba.ForEach(x => ih.AppendLine("  " + x + (++cnt == liba.Count ? "" : ",")));
            ih.AppendLine("]");
            ih.AppendLine("library " + ShortName);
            ih.AppendLine("{");
            using (new IDLHelperTab(ih))
            {
                // How do I know I'm importing stdole2??!
                // Forward declare all interfaces.
                ih.AppendLine("// Forward declare all types defined in this typelib");

                // Need to collect all dumpable interface names, in case we have dispinterfaces which don't have
                // top level interfaces.  In THIS case, we'd dump the dispinterface.
                var interfaceNames = Children.Aggregate<ITlibNode, ICollection<string>>(new HashSet<string>(),
                    (x, y) =>
                    {
                        if ((y as OWInterface) != null)
                        {
                            x.Add(y.ShortName);
                        }
                        return x;
                    });
                Children.FindAll(x => (x as OWInterface) != null || (x as OWDispInterface) != null).ForEach(x => ih.AppendLine(x.Name));
                Children.FindAll(x => x.DisplayAtTLBLevel(interfaceNames)).ForEach(x => { x.BuildIDLInto(ih); ih.AppendLine(""); });
            }
            ih.AppendLine("};");
        }
    }
}

