using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class ElemDesc
    {
        public ElemDesc(ELEMDESC desc)
        {
            tdesc = new TypeDesc(desc.tdesc);
            paramdesc = new ParamDesc(desc.desc.paramdesc);
        }

        public TypeDesc tdesc { get; }
        public ParamDesc paramdesc { get; }
    }
}
