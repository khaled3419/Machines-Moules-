using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using Microsoft.Maui;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Controls.Hosting;
using Microsoft.Maui.Hosting;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder.UseMauiApp<MoulesApp>();
        return builder.Build();
    }
}

public class MoulesApp : Application
{
    protected override Window CreateWindow(IActivationState activationState)
    {
        return new Window(new NavigationPage(new MoulesPage()));
    }
}

#if ANDROID
[Android.App.Application]
public class MainApplication : MauiApplication
{
    public MainApplication(IntPtr handle, Android.Runtime.JniHandleOwnership ownership)
        : base(handle, ownership) { }

    protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();
}

[Android.App.Activity(
    Theme = "@style/Maui.SplashTheme",
    MainLauncher = true,
    ConfigurationChanges = Android.Content.PM.ConfigChanges.ScreenSize |
                           Android.Content.PM.ConfigChanges.Orientation |
                           Android.Content.PM.ConfigChanges.UiMode |
                           Android.Content.PM.ConfigChanges.ScreenLayout |
                           Android.Content.PM.ConfigChanges.SmallestScreenSize |
                           Android.Content.PM.ConfigChanges.Density)]
public class MainActivity : MauiAppCompatActivity { }
#endif

public class Moule
{
    public string Code;
    public string Nom;
    public double W, H, V, Poids, Force;
}

public class Machine
{
    public string Type;
    public double Force, HMin, VMin, WMax, PoidsMax;
    public bool Complet;
}

public static class Donnees
{
    public static string Fichier = "/storage/emulated/0/Documents/Machines_Moules.xlsx";

    public static bool AccesComplet()
    {
#if ANDROID
        if (Android.OS.Build.VERSION.SdkInt >= Android.OS.BuildVersionCodes.R)
            return Android.OS.Environment.IsExternalStorageManager;
#endif
        return true;
    }

    public static void DemanderAcces()
    {
#if ANDROID
        var activity = Microsoft.Maui.ApplicationModel.Platform.CurrentActivity;
        if (activity == null) return;

        var intent = new Android.Content.Intent(
            Android.Provider.Settings.ActionManageAppAllFilesAccessPermission);
        intent.SetData(Android.Net.Uri.Parse(
            "package:" + Android.App.Application.Context.PackageName));
        activity.StartActivity(intent);
#endif
    }

    public static double Num(IXLCell c)
    {
        try { return c.GetDouble(); }
        catch { return 0; }
    }

    public static bool Nombre(string s, out double v)
    {
        return double.TryParse((s ?? "").Trim().Replace(',', '.'),
                   NumberStyles.Any, CultureInfo.InvariantCulture, out v)
               && v >= 0;
    }

    public static void Preparer()
    {
        if (File.Exists(Fichier)) return;

        using (var wb = new XLWorkbook())
        {
            var m = wb.Worksheets.Add("Moules");
            m.Cell(1, 1).Value = "CodeMoule";
            m.Cell(1, 2).Value = "NomMoule";
            m.Cell(1, 3).Value = "W";
            m.Cell(1, 4).Value = "H";
            m.Cell(1, 5).Value = "V";
            m.Cell(1, 6).Value = "Poids";
            m.Cell(1, 7).Value = "Force";

            var a = wb.Worksheets.Add("Machines");
            a.Cell(1, 1).Value = "TypeMachine";
            a.Cell(1, 2).Value = "ForceMachine";
            a.Cell(1, 3).Value = "MinH";
            a.Cell(1, 4).Value = "MinV";
            a.Cell(1, 5).Value = "MinW";
            a.Cell(1, 6).Value = "MaxW";
            a.Cell(1, 7).Value = "PoidsMax";

            wb.SaveAs(Fichier);
        }
    }

    public static List<Moule> LireMoules(XLWorkbook wb)
    {
        var liste = new List<Moule>();
        var ws = wb.Worksheet("Moules");

        foreach (var row in ws.RowsUsed())
        {
            if (row.RowNumber() == 1) continue;
            string code = row.Cell(1).GetString().Trim();
            if (code == "") continue;

            liste.Add(new Moule
            {
                Code = code,
                Nom = row.Cell(2).GetString().Trim(),
                W = Num(row.Cell(3)),
                H = Num(row.Cell(4)),
                V = Num(row.Cell(5)),
                Poids = Num(row.Cell(6)),
                Force = Num(row.Cell(7))
            });
        }
        return liste;
    }

    public static List<Machine> LireMachines(XLWorkbook wb)
    {
        var liste = new List<Machine>();
        var ws = wb.Worksheet("Machines");

        foreach (var row in ws.RowsUsed())
        {
            if (row.RowNumber() == 1) continue;
            string type = row.Cell(1).GetString().Trim();
            if (type == "") continue;

            var m = new Machine
            {
                Type = type,
                Force = Num(row.Cell(2)),
                HMin = Num(row.Cell(3)),
                VMin = Num(row.Cell(4)),
                WMax = Num(row.Cell(6)),
                PoidsMax = Num(row.Cell(7)),
                Complet = !row.Cell(2).IsEmpty() && !row.Cell(3).IsEmpty() &&
                          !row.Cell(4).IsEmpty() && !row.Cell(6).IsEmpty() &&
                          !row.Cell(7).IsEmpty()
            };
            liste.Add(m);
        }
        return liste;
    }

    public static Moule Chercher(List<Moule> moules, string recherche)
    {
        foreach (var m in moules)
        {
            if (m.Code.Equals(recherche, StringComparison.OrdinalIgnoreCase) ||
                m.Nom.Equals(recherche, StringComparison.OrdinalIgnoreCase))
                return m;
        }
        return null;
    }

    public static List<string> Verifier(Moule m, Machine x, double marge)
    {
        var raisons = new List<string>();
        double vMin = x.VMin - marge;
        double hMin = x.HMin - marge;

        if (m.W > x.WMax)
            raisons.Add("W moule " + m.W + " > W max " + x.WMax);
        if (m.V < vMin)
            raisons.Add("V moule " + m.V + " < V min " + vMin);
        if (m.H < hMin)
            raisons.Add("H moule " + m.H + " < H min " + hMin);
        if (m.Force > x.Force)
            raisons.Add("Force moule " + m.Force + " > Force machine " + x.Force);
        if (m.Poids > x.PoidsMax)
            raisons.Add("Poids moule " + m.Poids + " > Poids max " + x.PoidsMax);

        return raisons;
    }
}

public class MoulesPage : ContentPage
{
    Entry entRecherche, entMarge;
    Switch swIncomp;
    Label lblResultat;
    VerticalStackLayout formAjout;
    Entry fCode, fNom, fW, fH, fV, fPoids, fForce;
    Button btnAcces;

    static Button Bouton(string texte, EventHandler action)
    {
        var b = new Button { Text = texte };
        b.Clicked += action;
        return b;
    }

    static Entry Champ(string placeholder, bool numerique)
    {
        var e = new Entry { Placeholder = placeholder };
        if (numerique) e.Keyboard = Keyboard.Numeric;
        return e;
    }

    static Grid DeuxBoutons(Button a, Button b)
    {
        var g = new Grid { ColumnSpacing = 8 };
        g.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        g.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        Grid.SetColumn(a, 0);
        Grid.SetColumn(b, 1);
        g.Children.Add(a);
        g.Children.Add(b);
        return g;
    }

    public MoulesPage()
    {
        Title = "Moule / Machine";

        entRecherche = Champ("Code ou nom du moule", false);
        entMarge = Champ("Marge sur H min et V min (vide = aucune)", true);

        swIncomp = new Switch();
        var ligneSwitch = new HorizontalStackLayout { Spacing = 8 };
        ligneSwitch.Children.Add(swIncomp);
        ligneSwitch.Children.Add(new Label
        {
            Text = "Afficher aussi les machines non compatibles",
            VerticalOptions = LayoutOptions.Center
        });

        lblResultat = new Label { FontSize = 14 };

        fCode = Champ("Code du moule", false);
        fNom = Champ("Nom du moule", false);
        fW = Champ("W", true);
        fH = Champ("H", true);
        fV = Champ("V", true);
        fPoids = Champ("Poids (kg)", true);
        fForce = Champ("Force de verrouillage (T)", true);

        formAjout = new VerticalStackLayout { Spacing = 6, IsVisible = false };
        formAjout.Children.Add(fCode);
        formAjout.Children.Add(fNom);
        formAjout.Children.Add(fW);
        formAjout.Children.Add(fH);
        formAjout.Children.Add(fV);
        formAjout.Children.Add(fPoids);
        formAjout.Children.Add(fForce);
        formAjout.Children.Add(Bouton("💾 Enregistrer le moule", OnEnregistrer));

        var pile = new VerticalStackLayout { Padding = 16, Spacing = 10 };

        btnAcces = Bouton("🔓 Autoriser l'accès aux fichiers", OnAcces);
        btnAcces.IsVisible = !Donnees.AccesComplet();

        pile.Children.Add(btnAcces);
        pile.Children.Add(entRecherche);
        pile.Children.Add(entMarge);
        pile.Children.Add(ligneSwitch);
        pile.Children.Add(DeuxBoutons(
            Bouton("✅ Vérifier", OnVerifier),
            Bouton("🗑 Supprimer", OnSupprimer)));
        pile.Children.Add(DeuxBoutons(
            Bouton("📋 Moules", OnMoules),
            Bouton("🏭 Machines", OnMachines)));
        pile.Children.Add(Bouton("➕ Ajouter un moule", OnToggleForm));
        pile.Children.Add(formAjout);
        pile.Children.Add(lblResultat);

        Content = new ScrollView { Content = pile };
    }

    void Afficher(string texte) => lblResultat.Text = texte;

    void OnAcces(object sender, EventArgs e) => Donnees.DemanderAcces();

    bool VerifierAcces()
    {
        if (Donnees.AccesComplet())
        {
            btnAcces.IsVisible = false;
            return true;
        }
        btnAcces.IsVisible = true;
        Afficher("⚠️ Autorisez d'abord l'accès aux fichiers (bouton en haut), activez l'option pour cette application, puis revenez ici et réessayez.");
        return false;
    }

    bool LireMarge(out double marge)
    {
        marge = 0;
        string s = (entMarge.Text ?? "").Trim();
        if (s == "") return true;
        return Donnees.Nombre(s, out marge);
    }

    void OnVerifier(object sender, EventArgs e)
    {
        try
        {
            string recherche = (entRecherche.Text ?? "").Trim();
            if (recherche == "")
            {
                Afficher("❌ Entrez un code ou un nom de moule.");
                return;
            }

            double marge;
            if (!LireMarge(out marge))
            {
                Afficher("❌ Marge incorrecte.");
                return;
            }

            if (!VerifierAcces()) return;
            Donnees.Preparer();

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                var moules = Donnees.LireMoules(wb);
                var machines = Donnees.LireMachines(wb);
                Moule m = Donnees.Chercher(moules, recherche);

                if (m == null)
                {
                    Afficher("❌ Moule introuvable.");
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine("=== DONNEES DU MOULE ===");
                sb.AppendLine("Code : " + m.Code);
                sb.AppendLine("Nom : " + m.Nom);
                sb.AppendLine("W : " + m.W + " | H : " + m.H + " | V : " + m.V);
                sb.AppendLine("Poids : " + m.Poids + " kg | Force : " + m.Force + " T");
                sb.AppendLine("Marge : " + marge);
                sb.AppendLine();
                sb.AppendLine("=== MACHINES COMPATIBLES ===");

                var incompatibles = new StringBuilder();
                var incompletes = new List<string>();
                int nb = 0;

                foreach (var x in machines)
                {
                    if (!x.Complet)
                    {
                        incompletes.Add(x.Type);
                        continue;
                    }

                    var raisons = Donnees.Verifier(m, x, marge);

                    if (raisons.Count == 0)
                    {
                        nb++;
                        sb.AppendLine("✅ " + x.Type + " : le moule peut être monté sur cette machine");
                        sb.AppendLine("   Force " + x.Force + " T | W max " + x.WMax + " | H min " + x.HMin + " | V min " + x.VMin + " | Poids max " + x.PoidsMax);
                    }
                    else
                    {
                        incompatibles.AppendLine("❌ " + x.Type);
                        foreach (var r in raisons)
                            incompatibles.AppendLine("   - " + r);
                    }
                }

                if (nb == 0)
                    sb.AppendLine("❌ Aucune machine compatible.");

                if (swIncomp.IsToggled && incompatibles.Length > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("=== NON COMPATIBLES ===");
                    sb.Append(incompatibles.ToString());
                }

                if (incompletes.Count > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("⚠️ Non vérifiées (données manquantes dans Excel) : " + string.Join(", ", incompletes));
                }

                Afficher(sb.ToString());
            }
        }
        catch (Exception ex)
        {
            Afficher("❌ ERREUR : " + ex.Message);
        }
    }

    void OnToggleForm(object sender, EventArgs e) => formAjout.IsVisible = !formAjout.IsVisible;

    void OnEnregistrer(object sender, EventArgs e)
    {
        try
        {
            string code = (fCode.Text ?? "").Trim();
            string nom = (fNom.Text ?? "").Trim();

            if (code == "")
            {
                Afficher("❌ Le code du moule est obligatoire.");
                return;
            }

            double w, h, v, poids, force;
            if (!Donnees.Nombre(fW.Text, out w) || !Donnees.Nombre(fH.Text, out h) ||
                !Donnees.Nombre(fV.Text, out v) || !Donnees.Nombre(fPoids.Text, out poids) ||
                !Donnees.Nombre(fForce.Text, out force))
            {
                Afficher("❌ Valeurs numériques incorrectes.");
                return;
            }

            if (!VerifierAcces()) return;
            Donnees.Preparer();

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                var moules = Donnees.LireMoules(wb);
                if (Donnees.Chercher(moules, code) != null)
                {
                    Afficher("❌ Ce code existe déjà.");
                    return;
                }

                var ws = wb.Worksheet("Moules");
                var derniere = ws.LastRowUsed();
                int ligne = (derniere == null) ? 2 : derniere.RowNumber() + 1;

                ws.Cell(ligne, 1).Value = code;
                ws.Cell(ligne, 2).Value = nom;
                ws.Cell(ligne, 3).Value = w;
                ws.Cell(ligne, 4).Value = h;
                ws.Cell(ligne, 5).Value = v;
                ws.Cell(ligne, 6).Value = poids;
                ws.Cell(ligne, 7).Value = force;

                wb.Save();
            }

            fCode.Text = ""; fNom.Text = ""; fW.Text = ""; fH.Text = "";
            fV.Text = ""; fPoids.Text = ""; fForce.Text = "";
            formAjout.IsVisible = false;
            Afficher("✅ Moule " + code + " ajouté.");
        }
        catch (Exception ex)
        {
            Afficher("❌ ERREUR : " + ex.Message);
        }
    }

    async void OnSupprimer(object sender, EventArgs e)
    {
        try
        {
            string recherche = (entRecherche.Text ?? "").Trim();
            if (recherche == "")
            {
                Afficher("❌ Entrez le code ou le nom du moule à supprimer.");
                return;
            }

            if (!VerifierAcces()) return;
            Donnees.Preparer();

            int ligneTrouvee = -1;
            string codeTrouve = "";

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                var ws = wb.Worksheet("Moules");
                foreach (var row in ws.RowsUsed())
                {
                    if (row.RowNumber() == 1) continue;
                    string code = row.Cell(1).GetString().Trim();
                    string nom = row.Cell(2).GetString().Trim();

                    if (code.Equals(recherche, StringComparison.OrdinalIgnoreCase) ||
                        nom.Equals(recherche, StringComparison.OrdinalIgnoreCase))
                    {
                        ligneTrouvee = row.RowNumber();
                        codeTrouve = code;
                        break;
                    }
                }
            }

            if (ligneTrouvee == -1)
            {
                Afficher("❌ Moule introuvable.");
                return;
            }

            bool ok = await DisplayAlert("Supprimer", "Supprimer le moule " + codeTrouve + " ?", "Oui", "Non");
            if (!ok)
            {
                Afficher("Annulé.");
                return;
            }

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                wb.Worksheet("Moules").Row(ligneTrouvee).Delete();
                wb.Save();
            }

            Afficher("✅ Moule " + codeTrouve + " supprimé.");
        }
        catch (Exception ex)
        {
            Afficher("❌ ERREUR : " + ex.Message);
        }
    }

    void OnMoules(object sender, EventArgs e)
    {
        try
        {
            if (!VerifierAcces()) return;
            Donnees.Preparer();

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                var moules = Donnees.LireMoules(wb);
                var sb = new StringBuilder();
                sb.AppendLine("=== MOULES (" + moules.Count + ") ===");

                foreach (var m in moules)
                {
                    sb.AppendLine();
                    sb.AppendLine("• " + m.Code + " - " + m.Nom);
                    sb.AppendLine("  W " + m.W + " | H " + m.H + " | V " + m.V + " | Poids " + m.Poids + " kg | Force " + m.Force + " T");
                }

                Afficher(sb.ToString());
            }
        }
        catch (Exception ex)
        {
            Afficher("❌ ERREUR : " + ex.Message);
        }
    }

    void OnMachines(object sender, EventArgs e)
    {
        try
        {
            if (!VerifierAcces()) return;
            Donnees.Preparer();

            using (var wb = new XLWorkbook(Donnees.Fichier))
            {
                var machines = Donnees.LireMachines(wb);
                var sb = new StringBuilder();
                sb.AppendLine("=== MACHINES (" + machines.Count + ") ===");

                foreach (var x in machines)
                {
                    sb.AppendLine();
                    sb.AppendLine("• " + x.Type);

                    if (!x.Complet)
                    {
                        sb.AppendLine("  ⚠️ Données incomplètes dans Excel");
                        continue;
                    }

                    sb.AppendLine("  Force " + x.Force + " T | W max " + x.WMax + " | H min " + x.HMin + " | V min " + x.VMin + " | Poids max " + x.PoidsMax + " kg");
                }

                Afficher(sb.ToString());
            }
        }
        catch (Exception ex)
        {
            Afficher("❌ ERREUR : " + ex.Message);
        }
    }
}
