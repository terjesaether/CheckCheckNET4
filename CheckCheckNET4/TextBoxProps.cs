using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckCheckNET4
{

    static class TextBlockProps
    {
        public static event EventHandler<PropertyChangedEventArgs> StaticPropertyChanged;

        private static string _totNumLines;
        public static string TotNumLines
        {
            get { return _totNumLines; }
            set
            {
                if (value != _totNumLines)
                {
                    _totNumLines = value;

                    NotifyStaticPropertyChanged("TotNumLines");
                }
            }
        }

        private static bool _checkBoxIntroIsChecked;
        public static bool CheckBoxIntroIsChecked
        {
            get { return _checkBoxIntroIsChecked; }
            set
            {
                if (value != _checkBoxIntroIsChecked)
                {
                    _checkBoxIntroIsChecked = value;

                    NotifyStaticPropertyChanged("CheckBoxIntroIsChecked");
                }
            }
        }

        private static void NotifyStaticPropertyChanged(string propertyName)
        {
            if (StaticPropertyChanged != null)
                StaticPropertyChanged(null, new PropertyChangedEventArgs(propertyName));
        }
    }


}
