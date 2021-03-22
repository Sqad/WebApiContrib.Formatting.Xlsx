using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace SQAD.XlsxExportImport.Base.Models
{
    public class KeyedCollectionBase<T> : KeyedCollection<string, T>
    {
        public ICollection<string> Keys
        {
            get
            {
                if (Dictionary != null)
                {
                    return Dictionary.Keys;
                }
                else
                {
                    return new Collection<string>(this.Select(GetKeyForItem).ToArray());
                }
            }
        }

        protected override string GetKeyForItem(T item)
        {
            throw new NotImplementedException();
        }
    }
}
