using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class KeyedCollectionBase<T> : KeyedCollection<string, T>
    {
        public ICollection<string> Keys
        {
            get
            {
                if (this.Dictionary != null)
                {
                    return this.Dictionary.Keys;
                }
                else
                {
                    return new Collection<string>(this.Select(this.GetKeyForItem).ToArray());
                }
            }
        }

        protected override string GetKeyForItem(T item)
        {
            throw new NotImplementedException();
        }
    }
}
