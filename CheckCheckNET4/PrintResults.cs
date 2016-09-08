using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CheckCheckNET4
{
    public class PrintResult
    {


        public static int printResultByEpisode(List<Episode> episodeList, ListBox resultPanel, CheckBox checkBoxNoIntro)
        {
            int totNumLines = 0;
            int counter = 0;

            foreach (var item in episodeList)
            {
                if (item.roleNames.Count > 0 || item.Mengder.Count > 0)
                {

                    var container = new StackPanel { Margin = new Thickness(5, 5, 0, 5) };

                    var mainText = new TextBox { FontSize = 15, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, -30), BorderThickness = new Thickness(0), Padding = new Thickness(4, 4, 4, 4) };

                    var textBlockHeading = new TextBlock { FontSize = 18, FontWeight = FontWeights.Bold, Foreground = Brushes.WhiteSmoke, Background = Brushes.DarkBlue, Height = 24, Padding = new Thickness(4, 0, 4, 4) };

                    var textBlockMengde = new TextBlock { FontSize = 16, Text = "Mengder: ", Foreground = Brushes.Red, Margin = new Thickness(0, 0, 0, 10) };

                    var removeButton = new Button { Content = "Fjern", Width = 80, Height = 24, HorizontalAlignment = HorizontalAlignment.Left };

                    removeButton.Click += btn_Click;

                    TimeSpan span = new TimeSpan();

                    int totNumLinesPrEpisode = 0;

                    string warningText = "";
                    string daysText = " dager.";
                    string deliveryText = "";
                    //string resultTemp = "";
                    string mengdeTemp = "";
                    string mengdeText = "";
                    string resultText = "";


                    if (item.deliveryDate.ToString() != "")
                    {
                        try
                        {
                            DateTime deliveryDate = Convert.ToDateTime(item.deliveryDate.ToString());
                            span = deliveryDate.Subtract(DateTime.Now);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Åh heisann! Virker som det er et Excelark som ikke er i orden. Kan det være " + item.seriesName.ToString() + ", ep " + item.episodeNumber.ToString() + " ?");
                            break;
                        }

                        if (span.Days < 5)
                        {
                            warningText = " På tide å KLIPPE!!";
                        }
                        deliveryText = ". Skal leveres om " + span.Days.ToString() + daysText + warningText;
                    }


                    if (span.Days == 1)
                    {
                        daysText = " dag.";
                    }

                    // TextBlockHeading
                    textBlockHeading.Text = string.Format("{0} - Episode {1} {2}", item.seriesName.ToUpper(), item.episodeNumber, deliveryText);


                    string resultTemp = "";

                    resultTemp += Environment.NewLine;

                    for (int role = 0; role < item.roleNames.Count; role += 1)
                    {

                        if (!(checkBoxNoIntro.IsChecked.Value && (item.roleNames[role].searchName.ToString().ToLower() == "intro" || item.roleNames[role].searchName.ToString().ToLower() == "outro" || item.roleNames[role].searchName.ToString().ToLower() == "plakat")))
                        {
                            resultTemp += item.roleNames[role].searchName.ToString() + ": " + item.roleNames[role].numOfLines.ToString() + ", ";

                            // Regner ut antall replikker
                            totNumLinesPrEpisode += Convert.ToInt32(item.roleNames[role].numOfLines);

                        }
                    }



                    // Fyller opp mengde-info:
                    for (int m = 0; m < item.Mengder.Count; m++)
                    {
                        mengdeTemp += item.Mengder[m].MengdeRoleNames.ToString() + " (" + item.Mengder[m].MengdeNumOfLines.ToString() + "), ";

                    }
                    // Tar bort komma på slutten:
                    if (resultTemp.Length > 2)
                    {
                        resultText = resultTemp.Substring(0, resultTemp.Length - 2);
                    }
                    if (mengdeTemp.Length > 2)
                    {
                        mengdeText = mengdeTemp.Substring(0, mengdeTemp.Length - 2);
                    }

                    resultText += " == Totalt " + totNumLinesPrEpisode.ToString() + " rep.";

                    if (totNumLinesPrEpisode > 0)
                    {
                        // Sjekker om få replikker
                        if (totNumLinesPrEpisode < 3)
                        {
                            resultText += " ..kanskje vi har en pickup her?";
                        }

                        resultText += "\r\n\r\n";


                        // Slår sammen hele WrapPanelet med TextBlock'ene:
                        //mainText.Text = resultText;
                        mainText.Text = resultText;
                        textBlockMengde.Text += mengdeText;
                        container.Children.Add(textBlockHeading);
                        container.Children.Add(mainText);

                        if (mengdeTemp.Length > 0)
                        {
                            container.Children.Add(textBlockMengde);
                        }

                        container.Children.Add(removeButton);

                        resultPanel.Items.Add(container);

                    }
                    totNumLines += totNumLinesPrEpisode;
                }
                counter++;

            }
            return totNumLines;
        }

        public static void btn_Click(object sender, EventArgs e)
        {
            var objButton = (Button)sender;
            StackPanel objWrap = (StackPanel)objButton.Parent;
            //StackPanel objStackPanel = (StackPanel)objWrap.Parent;
            var objContainer = (ListBox)objWrap.Parent;

            //MessageBox.Show(objWrap.Name.ToString() + " ble fjernet.");

            //objStackPanel.Children.Remove(objWrap);
            objContainer.Items.Remove(objWrap);


        }



    }
}
