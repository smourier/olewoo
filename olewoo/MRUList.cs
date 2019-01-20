using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace olewoo_cs
{
    class MRUList
    {
        const int MAXDATAITEMS = 9;

        RegistryKey _mrufiles;
        LinkedList<string> _data;
        Dictionary<string, bool> _uqitems;

        public MRUList(string path)
        {
            RegistryKey hkcu = Registry.CurrentUser;
            _mrufiles = GetRegKey(path, hkcu);
            Refresh();
        }

        public string[] Items => _data.ToArray();

        public void Clear()
        {
            _data = new LinkedList<string>();
            _uqitems = new Dictionary<string, bool>(StringComparer.CurrentCultureIgnoreCase);
        }

        public void AddItem(string path)
        {
            if (!_uqitems.ContainsKey(path))
            {
                while (_data.Count >= MAXDATAITEMS)
                {
                    string last = _data.Last.Value;
                    _uqitems.Remove(last);
                    _data.RemoveLast();
                }
                _data.AddFirst(path);
                _uqitems[path] = true;
            }
        }

        private void Refresh()
        {
            Clear();
            for (int i = 0; i < 10; ++i)
            {
                object o = _mrufiles.GetValue("MruFile" + i);
                string s = o as string;
                if (string.IsNullOrEmpty(s)) return;
                _data.AddLast(s);
                _uqitems[s] = true;
            }
        }

        public void Flush()
        {
            int mval = _data.Count;
            string[]data = _data.ToArray();
            for (int i = 0; i < 10; ++i)
            {
                string keyname = "MruFile" + i;
                if (i >= mval)
                {
                    _mrufiles.SetValue(keyname, "");
                }
                else
                {
                    _mrufiles.SetValue(keyname, data[i]);
                }
            }
        }

        private RegistryKey GetRegKey(string key, RegistryKey basekey)
        {
            RegistryKey nkey = basekey.OpenSubKey(key, true);
            if (nkey == null) nkey = basekey.CreateSubKey(key);
            return nkey;
        }
    }
}
