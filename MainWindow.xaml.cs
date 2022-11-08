using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;


namespace CRM
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {  public List<Client> ListeClient { get; set; }
        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook classeur;
        private Worksheet feuille;
        private Microsoft.Office.Interop.Word.Application word;
        //sur word on travaille sur des objets document
        Document source;
        private Microsoft.Office.Interop.Outlook.Application outlook;
        private MailItem mail;

        public MainWindow()
        {
            InitializeComponent();
            ListeClient = new List<Client>();
            excel = new Microsoft.Office.Interop.Excel.Application();
            outlook = new Microsoft.Office.Interop.Outlook.Application();
            try
            {
                Navigateur.Source = new System.Uri(@"http://www.google.fr");

            }

            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message); 
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Ajouter_Client window = new Ajouter_Client();
            //Méthode bloquante
            window.ShowDialog();
           Client courantClient= window.Client;

            ListeClient.Add(courantClient);
          
            NomDataGrid.ItemsSource = null;
            NomDataGrid.ItemsSource = ListeClient;

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Fichier Excel | *.xlsx |Fichier CSV | * .csv";
            sfd.FilterIndex = 1;
            if (sfd.ShowDialog() == true)
            {
                bool fileExists = File.Exists(sfd.FileName);

                if (!fileExists || classeur != null)
                {
                    classeur = excel.Workbooks.Add();

                }
                else
                {
                    classeur = excel.Workbooks.Open(sfd.FileName);
                }

                feuille = classeur.ActiveSheet;

                int ligne = 1;
                //On écrit la liste dans un fichier Excel

                foreach ( Client client in ListeClient)// on parcours toute la liste client

                {
                    feuille.Cells[ligne, 1] = client.Id;
                    feuille.Cells[ligne, 2] = client.Prenom;
                    feuille.Cells[ligne, 3] = client.Nom;
                    feuille.Cells[ligne, 4] = client.Mail;
                    feuille.Cells[ligne, 5] = client.Tel;
                    feuille.Cells[ligne, 6] = client.Societe;
                    ligne++;

                }
                excel.DisplayAlerts = false;//Pour éviter les messages d'alerte fichier existant
                classeur.SaveAs(sfd.FileName);
                classeur.Close();
                excel.DisplayAlerts = true;//On réactive
            }

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
           
            ofd.Filter = "Fichier Excel | *.xlsx";
                     
           
            //Si un fichier a été choisi
            if (ofd.ShowDialog() == true)
               //Ouvre le classeur
                
            {
                classeur = excel.Workbooks.Open(ofd.FileName);

               // On selectione la feuille active

                feuille = classeur.ActiveSheet;

                int ligne = 1;

                //On lit le fichier et ajoute le client
                //tant que l'on en trouve pas une chaine vide on ajoute les clients
                while(feuille.Cells[ligne, 1].Text!=string.Empty)
                {
                    //Recoupere les clients

                    Client newClient = new Client(

                    int.Parse(feuille.Cells[ligne, 1].Text),
                      feuille.Cells[ligne, 2].Text,
                      feuille.Cells[ligne, 3].Text,
                      feuille.Cells[ligne, 4].Text,
                      feuille.Cells[ligne, 5].Text,
                      feuille.Cells[ligne, 6].Text
                      
                    );

                    ligne++;
                    //Ajouter à la liste  client
                    ListeClient.Add(newClient);

                }
                      NomDataGrid.ItemsSource = ListeClient;



        }

        }

        private void NomDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int index = NomDataGrid.SelectedIndex;
            if (index !=-1)
            {
                Client clientModifier = (Client)NomDataGrid.Items[index];
                Ajouter_Client ajoutClient = new Ajouter_Client();
                ajoutClient.Client = clientModifier;
                ajoutClient.UpdateWPF();
                ajoutClient.ShowDialog();
                ListeClient.Remove(clientModifier);
                ListeClient.Insert(index, ajoutClient.Client);
                
                NomDataGrid.ItemsSource = null;
                NomDataGrid.ItemsSource = ListeClient;
            }


        }
        //Bouton de modification mailing
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            word = new Microsoft.Office.Interop.Word.Application();
            source = word.Documents.Open(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.docx");
            word.Visible = true;

            //On stop le code jusqu'ai modification 
            MessageBox.Show(this,
                "Modifier le document Word puis OK",
                "Modification",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            //Bloquer par l'appel precedent
            //MessageBox.Show("On passe dans le code après");
            word.Quit();
            
                word = new Microsoft.Office.Interop.Word.Application();
                source = word.Documents.Open(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.docx");
                word.Visible = false;
            
            //Document vierge
            Document destination = word.Documents.Add();
            Microsoft.Office.Interop.Word.Range selection = source.Content;
           // On copie la selection
            selection.Copy();
            //on cole le document word
            destination.Content.PasteSpecial(WdPasteOptions.wdKeepSourceFormatting);
            destination.SaveAs2(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.html", WdSaveFormat.wdFormatHTML);

            try
            {
                Navigateur.Source = new System.Uri(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.html");

            }

            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
           
            foreach (Client client in ListeClient)
            {
                mail = outlook.CreateItem(OlItemType.olMailItem);
                mail.To = client.Nom + " " + client.Prenom + " <" + client.Mail + ">";
                mail.Subject = "Rapport client: " + client.Societe;
                mail.Attachments.Add(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.docx");
                mail.HTMLBody = System.IO.File.ReadAllText(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.html");

#if DEBUG
                //On affiche le mail
                mail.Display(true);

#else           
               //On souhaite envoyer
                mail.Send();
#endif
            }
            
                    

        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            word = new Microsoft.Office.Interop.Word.Application();
            source = word.Documents.Open(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.docx");
            Document destination = word.Documents.Add();
            Microsoft.Office.Interop.Word.Range selection = source.Content;
            // On copie la selection
            selection.Copy();
            //on cole le document word
            destination.Content.PasteSpecial(WdPasteOptions.wdKeepSourceFormatting);
            destination.SaveAs2(@"C:\Users\admin\Documents\Microsoft_operabilite\mailling.pdf", WdSaveFormat.wdFormatPDF);


        }
    }
}
