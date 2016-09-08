using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace CheckCheckNET4
{
    class Productions
    {
    }

    // Liste over alle produksjonene, hver DataTable er en forside
    public class NordubbProductions
    {
        private List<excelFrontPage> _productions = new List<excelFrontPage>();

        public List<excelFrontPage> productions
        {
            get { return _productions; }
            set { _productions = value; }
        }

    }

    // En serie / forside består av et navn og en liste med Produksjoner
    public class excelFrontPage
    {
        private List<string> _numEpisodes = new List<string>();

        public string seriesName { get; set; }
        public string excelFileName { get; set; }
        public DataTable frontPageDataTable { get; set; }

        public List<string> numEpisodesList
        {
            get { return _numEpisodes; }
            set { _numEpisodes = value; }
        }

        public string trimFilename(string filename, string dubToolDir)
        {
            string t = filename.Substring(dubToolDir.Length);
            t = t.Substring(0, t.Length - 4);
            return t;
        }
    }

    public class Episode
    {
        public string episodeNumber { get; set; }
        public string seriesName { get; set; }
        public List<SearchNameAndNumOfLines> roleNames { get; set; }
        public List<Mengder> Mengder { get; set; }
        public string deliveryDate { get; set; }

        // Constructor
        public Episode()
        {
            roleNames = new List<SearchNameAndNumOfLines>();
            Mengder = new List<Mengder>();
        }
    }

    public class SearchNameAndNumOfLines
    {
        public string searchName { get; set; }
        public string numOfLines { get; set; }
        public string totalNumOfLines { get; set; }
    }

    public class Mengder
    {
        public string MengdeRoleNames { get; set; }
        public string MengdeNumOfLines { get; set; }
    }
}
