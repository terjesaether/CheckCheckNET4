using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using static CheckCheckNET4.MainWindow;

namespace CheckCheckNET4
{
    public class Calculations
    {
        // En Episode-class består av Episodenummer og Role med antall linjer
        public static List<Episode> calculateSearchResultsByEpisode(DataTable dt, string searchString)
        {
            string currEpsLines = "",
            linesTotal = "",
            linesDone,
            roleCell,
            searchCell = "",
            currentRoleName = "",
            currentRoleLineNumbersString = "";

            string[] linesArray;


            // Dette er en liste med Episoder, Episoder med en liste over rollene i den episoden
            var episodeList = new List<Episode>();


            // Går gjennom alle kolonnene: [rad][kolonne]
            for (int epColumn = 4; epColumn <= 16; epColumn++)
            {
                int rowCompensation = 0;

                Episode episode = new Episode(); // Oppretter et ny episodeobjekt

                episode.seriesName = dt.Rows[0][0].ToString();

                // Sjekker om manuset er i orden: sjekker hvor tittelen står:
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (string.IsNullOrEmpty(episode.seriesName))
                    {
                        rowCompensation++;
                        episode.seriesName = dt.Rows[rowCompensation][0].ToString();
                    }
                    else
                    {
                        break;
                    }
                }
                // Legger til epnr:
                episode.episodeNumber = dt.Rows[2 + rowCompensation][epColumn].ToString();

                episode.deliveryDate = dt.Rows[4 + rowCompensation][epColumn].ToString();


                // Scanner nedover
                for (int row = (5 + rowCompensation); row < dt.Rows.Count; row++)
                {
                    // Hvis kolonnen ikke er tom..:
                    if (dt.Rows[row][epColumn] != null)
                    {

                        var roleNamesAndNumLines = new SearchNameAndNumOfLines();


                        if (GlobalVariables.searchType == "actor")
                        {
                            searchCell = dt.Rows[row][3].ToString().ToLower();
                        }
                        else
                        {
                            searchCell = dt.Rows[row][1].ToString().ToLower();
                        }



                        // Hvis skuespillercellen inneholder søkestrengen..
                        if (searchCell.Contains(searchString))
                        {
                            // Deler opp replikk-cellen
                            currEpsLines = dt.Rows[row][epColumn].ToString().ToLower();
                            linesArray = currEpsLines.Split(new char[] { '/' }, StringSplitOptions.None);

                            if (linesArray.Length == 2)
                            {
                                linesDone = linesArray[0];
                                linesTotal = linesArray[1];

                                // Sjekker om det er mer enn 0 replikker igjen:
                                if (Convert.ToInt32(linesTotal) - Convert.ToInt32(linesDone) > 0)
                                {
                                    // Finner episodenummer og rollenavnet i kolonne:
                                    if (GlobalVariables.searchType == "actor")
                                    {
                                        currentRoleName = dt.Rows[row][1].ToString().ToUpper();
                                    }
                                    else
                                    {
                                        currentRoleName = dt.Rows[row][3].ToString().ToUpper();
                                    }

                                    currentRoleLineNumbersString = (Convert.ToInt32(linesTotal) - Convert.ToInt32(linesDone)).ToString();

                                    roleNamesAndNumLines.searchName = currentRoleName;
                                    roleNamesAndNumLines.numOfLines = currentRoleLineNumbersString;
                                    roleNamesAndNumLines.totalNumOfLines = linesTotal;


                                    // Legger til rollenavnet i episoden:
                                    episode.roleNames.Add(roleNamesAndNumLines);

                                }
                            }
                        }

                        roleCell = dt.Rows[row][1].ToString().ToLower();


                        if (roleCell.Contains("mengde") || roleCell.Contains("folk") || roleCell.Contains("publikum") || roleCell.Contains("unger") || roleCell.Contains("alle") || roleCell.Contains("menn") || roleCell.Contains("damer") || roleCell.Contains("gjester"))
                        {
                            // Deler opp replikk-cellen
                            currEpsLines = dt.Rows[row][epColumn].ToString().ToLower();
                            linesArray = currEpsLines.Split(new char[] { '/' }, StringSplitOptions.None);

                            if (linesArray.Length == 2)
                            {
                                linesDone = linesArray[0];
                                linesTotal = linesArray[1];

                                // Sjekker om det er mer enn 0 replikker igjen:
                                if (Convert.ToInt32(linesTotal) - Convert.ToInt32(linesDone) > 0)
                                {
                                    // Finner episodenummer i kolonne:
                                    currentRoleName = dt.Rows[row][1].ToString().ToUpper();
                                    currentRoleLineNumbersString = (Convert.ToInt32(linesTotal) - Convert.ToInt32(linesDone)).ToString();

                                    Mengder m = new Mengder
                                    {
                                        MengdeNumOfLines = currentRoleLineNumbersString,
                                        MengdeRoleNames = currentRoleName
                                    };

                                    episode.Mengder.Add(m);

                                }
                            }
                        }

                    }
                }
                // Testplassering:
                episodeList.Add(episode);
            }

            return episodeList;
        }

        // REGNER UT ** EN ** SERIE
        public static void calculateByOneEpisode(DataTable dt, string searchString, ListBox flowLayoutPanel1, CheckBox chckIntro, string linesPrEpisode, TextBlock lblTotalNumLines)
        {
            //MyVariables.searchType

            List<Episode> episodeList = Calculations.calculateSearchResultsByEpisode(dt, searchString);

            int totNumLines = PrintResult.printResultByEpisode(episodeList, flowLayoutPanel1, chckIntro);
            decimal t = decimal.Round((Convert.ToDecimal(totNumLines) / Convert.ToInt32(linesPrEpisode)), 2);

            lblTotalNumLines.Text = "TOTALT antall replikker: " + Convert.ToString(totNumLines) + ". Det vil ta " + t + " timer å dubbe ferdig.";



        }

        public static void calculateSelectedEpisodes(List<DataTable> ldt, string searchString, ListBox flowLayoutPanel1, CheckBox chckIntro, string linesPrEpisode, TextBlock lblTotalNumLines)
        {
            foreach (var item in ldt)
            {
                calculateByOneEpisode(item, searchString, flowLayoutPanel1, chckIntro, linesPrEpisode, lblTotalNumLines);
            }
        }

        // Regner ut ** ALLE ** seriene 
        public static void calculateAllSeriesAndEpisodes(NordubbProductions productions, string searchString, ListBox flowLayoutPanel1, CheckBox chckIntro, TextBlock lblTotalNumLines)
        {
            int totNumLines = 0;
            foreach (var item in productions.productions)
            {
                //string seriesName = "";

                List<Episode> episodeList = calculateSearchResultsByEpisode(item.frontPageDataTable, searchString);

                if (episodeList.Count > 0)
                {
                    //seriesName = item.seriesName.ToString();
                    totNumLines = totNumLines + PrintResult.printResultByEpisode(episodeList, flowLayoutPanel1, chckIntro);
                }

                // Må gi Utskriften en mulighet til å skrive ut navn på eps

            };

            decimal t = decimal.Round((Convert.ToDecimal(totNumLines) / 90), 2);
            lblTotalNumLines.Text = "TOTALT antall replikker: " + totNumLines.ToString() + ".  Det vil ta " + t + " timer å dubbe ferdig.";

            //TextBlockProps.TotNumLines = "TOTALT antall replikker: " + totNumLines.ToString() + ".  Det vil ta " + t + " timer å dubbe ferdig.";
        }


    }


}
