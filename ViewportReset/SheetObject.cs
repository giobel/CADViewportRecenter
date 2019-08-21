

namespace AttributeUpdater
{
    class SheetObject
    {
        public string sheetName { get; private set; }
        public XYZ viewCentre { get; private set; }
        public double angleToNorth { get; private set; }
        public XYZ viewportCentre { get; private set; }
        public double viewportWidth { get; private set; }
        public double viewportHeight { get; private set; }
        public string xrefName { get; private set; }

        public SheetObject(string SheetName, XYZ ViewCentre, double AngleToNorth, XYZ ViewportCentre, double ViewportWidth, double ViewportHeight, string XrefName)
        {
            sheetName = SheetName;
            viewCentre = ViewCentre;
            angleToNorth = AngleToNorth;
            viewportCentre = ViewportCentre;
            viewportWidth = ViewportWidth;
            viewportHeight = ViewportHeight;
            xrefName = XrefName;

        }
    }
}
