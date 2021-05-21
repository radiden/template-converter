namespace templater
{
    public class BedInfo
    {
        public string Tekst;
        public string Model;
        public string RozmiarRamy;
        public string TekstRozmiar;
        public string KolorRamy;
        public string TekstKolor;
        public string KolorSzuflady;
        public string TekstSzuflada;
    }

    public class FinalBedInfo
    {
        public string Model { get; set; }
        public string Template { get; set; }
    }
}