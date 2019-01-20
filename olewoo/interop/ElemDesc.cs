using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class ElemDesc
    {
        private ELEMDESC _desc;

        public ElemDesc(ELEMDESC desc)
        {
            _desc = desc;
        }

        public TypeDesc tdesc => new TypeDesc(_desc.tdesc);
        public ParamDesc paramdesc => new ParamDesc(_desc.desc.paramdesc);
    }
}
