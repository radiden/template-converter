using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FastExcel;

namespace templater
{
    class Program
    {
        static void Main(string[] args)
        {
            var input = new FileInfo(args[0]);
            // var input = new FileInfo("/home/rai/coding/cs/templater/DaneOpisy.xlsx");

            if (!input.Exists)
            {
                throw new FileNotFoundException($"Nie mogę znaleźć pliku {input.Name}!");
            }
            
            using (var reader = new FastExcel.FastExcel(input, true))
            {
                var worksheet = reader.Read(1);
                
                List<string> teksty = Utils.GetCellsAndNumber<string>("TekstReplaced", worksheet);
                List<string> modele = Utils.GetCellsAndNumber<string>("Model", worksheet);
                List<string> kolory = Utils.GetCellsAndNumber<string>("KolorRamy", worksheet);
                List<string> kolorySzuflad = Utils.GetCellsAndNumber<string>("KolorSzuflady", worksheet);
                List<string> rozmiary = Utils.GetCellsAndNumber<string>("RozmiarRamy", worksheet);
                List<string> tekstKolor = Utils.GetCellsAndNumber<string>("ZdanieKolorReplaced", worksheet);
                List<string> tekstSzuflada = Utils.GetCellsAndNumber<string>("ZdanieSzufladaReplaced", worksheet);
                List<string> tekstRozmiar = Utils.GetCellsAndNumber<string>("ZdanieRozmiarReplaced", worksheet);

                List<BedInfo> Lozka = new();
                
                for (var i = 0; i < teksty.Count; i++)
                {
                    BedInfo lozko = new();
                    lozko.Tekst = teksty[i];
                    lozko.Model = modele[i];
                    lozko.KolorRamy = kolory[i];
                    lozko.KolorSzuflady = kolorySzuflad[i];
                    lozko.RozmiarRamy = rozmiary[i];
                    lozko.TekstKolor = tekstKolor[i];
                    lozko.TekstRozmiar = tekstRozmiar[i];
                    lozko.TekstSzuflada = tekstSzuflada[i];
                    Lozka.Add(lozko);
                }

                var grouped = Lozka.GroupBy(l => l.Model);
                StringBuilder sb = new();
                var templateName = "out.xlsx";
                var outName = templateName;
                var iter = 0;
                while (File.Exists(outName))
                {
                    iter++;
                    var split = templateName.Split("."); 
                    outName = split[0] + iter + "." + split[1];
                }
                using (var output = new FastExcel.FastExcel(new FileInfo("template.xlsx"), new FileInfo(outName)))
                {
                    List<FinalBedInfo> finalBeds = new();
                    foreach (var group in grouped)
                    {
                        var currentTekst = group.FirstOrDefault().Tekst;
                        List<string> tekstLines = new();
                        tekstLines.AddRange(currentTekst.Split("\n"));
                        currentTekst = ReconstructTekstWithParagraphs(tekstLines);
                        // Console.WriteLine($" === {group.Key} === ");

                        // Console.WriteLine("= Ramy =");
                        sb.Append("\n#SWITCH {CECHA: Lóżka, wymiary}\n@DEFAULT:\n\n");
                        var ramaGroups = group.GroupBy(l => l.RozmiarRamy);
                        foreach (var ramaGroup in ramaGroups)
                        {
                            if (CheckNull(ramaGroup.Key))
                            {
                                Console.WriteLine($"Pomijam ramę NULL w grupie {group.Key}");
                                continue;
                            }
                            // sb.Append("\n#IF {CECHA:Lóżka, wymiary} = " + ramaGroup.Key + "\n{\n" +
                            //           ramaGroup.FirstOrDefault(l => l.RozmiarRamy == ramaGroup.Key).TekstRozmiar + "\n}");
                            sb.Append($"@VALUE: {ramaGroup.Key}\n{ramaGroup.FirstOrDefault(l => l.RozmiarRamy == ramaGroup.Key).TekstRozmiar}\n");
                            // Console.WriteLine("#IF {CECHA:Lóżka, wymiary} = " + ramaGroup.Key + "\n{\n" +
                            // ramaGroup.FirstOrDefault(l => l.RozmiarRamy == ramaGroup.Key).TekstRozmiar + "\n}");
                        }
                        
                        sb.Append("\n#ENDSWITCH\n");
                        currentTekst = currentTekst.Replace("#ROZMIAR", sb.ToString());
                        sb.Clear();

                        // Console.WriteLine("= Kolory ram =");
                        sb.Append("\n#SWITCH {WARIANT: Kolor ramy}\n@DEFAULT:\n\n");
                        var kolorGroups = group.GroupBy(l => l.KolorRamy);
                        foreach (var kolorGroup in kolorGroups)
                        {
                            if (CheckNull(kolorGroup.Key))
                            {
                                Console.WriteLine($"Pomijam kolor ramy NULL w grupie {group.Key}");
                                continue;
                            }
                            // sb.Append("\n#IF {WARIANTY:Kolor ramy} = " + kolorGroup.Key + "\n" +
                            //                   kolorGroup.FirstOrDefault(l => l.KolorRamy == kolorGroup.Key).TekstKolor + "\n#ENDIF");
                            sb.Append($"@VALUE: {kolorGroup.Key}\n{kolorGroup.FirstOrDefault(l => l.KolorRamy == kolorGroup.Key).TekstKolor}\n");
                            // Console.WriteLine("#IF {WARIANTY:Kolor ramy} = " + kolorGroup.Key + "\n" +
                            // kolorGroup.FirstOrDefault(l => l.KolorRamy == kolorGroup.Key).TekstKolor + "\n#ENDIF");
                        }
                        
                        sb.Append("\n#ENDSWITCH\n");
                        currentTekst = currentTekst.Replace("#KOLORRAMY", sb.ToString());
                        sb.Clear();

                        // Console.WriteLine("= Kolory szuflad =");
                        sb.Append("\n#SWITCH {WARIANT: Kolor szuflady}\n@DEFAULT:\n\n");
                        var szufladaKolorGroups = group.GroupBy(l => l.KolorSzuflady);
                        foreach (var szufladaKolorGroup in szufladaKolorGroups)
                        {
                            if (CheckNull(szufladaKolorGroup.Key))
                            {
                                Console.WriteLine($"Pomijam kolor szuflady NULL w grupie {group.Key}");
                                continue;
                            }
                            // sb.Append("\n#IF {WARIANTY:Kolor szuflady} = " + szufladaKolorGroup.Key + "\n" +
                            //                   szufladaKolorGroup.FirstOrDefault(l => l.KolorSzuflady == szufladaKolorGroup.Key).TekstSzuflada + "\n#ENDIF");
                            sb.Append($"@VALUE: {szufladaKolorGroup.Key}\n{szufladaKolorGroup.FirstOrDefault(l => l.KolorSzuflady == szufladaKolorGroup.Key).TekstSzuflada}\n");
                            // Console.WriteLine("#IF {WARIANTY:Kolor szuflady} = " + szufladaKolorGroup.Key + "\n" +
                            // szufladaKolorGroup.FirstOrDefault(l => l.KolorSzuflady == szufladaKolorGroup.Key).TekstSzuflada + "\n#ENDIF");
                        }

                        sb.Append("\n#ENDSWITCH\n");
                        currentTekst = currentTekst.Replace("#KOLORSZUFLADY", sb.ToString());
                        sb.Clear();
                        // Console.WriteLine(currentTekst);

                        var currentBed = new FinalBedInfo
                        {
                            Model = group.Key,
                            Template = currentTekst
                        };

                        finalBeds.Add(currentBed);
                    }
                    output.Write(finalBeds, "Sheet1");
                    Console.WriteLine($"Zapisano do {outName}!");
                }
            }
        }

        private static bool CheckNull(string str)
        {
            return (str == "NULL" || str == "");
        }

        private static string ReconstructTekstWithParagraphs(List<string> list)
        {
            StringBuilder sb = new();
            foreach (var line in list)
            {
                sb.Append("<p>\n" + line + "</p>\n");
            }

            return sb.ToString();
        }
    }
}

// #ROZMIAR  #KOLORRAMY #KOLORSZUFLADY