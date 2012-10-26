#region Namespaces
using System;
using System.Diagnostics;
using Autodesk.Revit.DB;
#endregion // Namespaces

namespace ObjExport
{
  class Util
  {
    /// <summary>
    /// Return an English plural suffix 's' or 
    /// nothing for the given number of items.
    /// </summary>
    public static string PluralSuffix( int n )
    {
      return 1 == n ? "" : "s";
    }

    /// <summary>
    /// Return a string for a real number
    /// formatted to two decimal places.
    /// </summary>
    public static string RealString( double a )
    {
      return a.ToString( "0.##" );
    }

    /// <summary>
    /// Return a string for an XYZ point
    /// or vector with its coordinates
    /// formatted to two decimal places.
    /// </summary>
    public static string PointString( XYZ p )
    {
      return string.Format( "({0},{1},{2})",
        RealString( p.X ),
        RealString( p.Y ),
        RealString( p.Z ) );
    }

    /// <summary>
    /// Return a string describing the given element:
    /// .NET type name, category name, family and 
    /// symbol name for a family instance, element id 
    /// and element name.
    /// </summary>
    public static string ElementDescription( Element e )
    {
      if( null == e )
      {
        return "<null>";
      }

      // For a wall, the element name equals the
      // wall type name, which is equivalent to the
      // family name ...

      FamilyInstance fi = e as FamilyInstance;

      string typeName = e.GetType().Name;

      string categoryName = ( null == e.Category )
        ? string.Empty
        : e.Category.Name + " ";

      string familyName = ( null == fi )
        ? string.Empty
        : fi.Symbol.Family.Name + " ";

      string symbolName = ( null == fi
        || e.Name.Equals( fi.Symbol.Name ) )
          ? string.Empty
          : fi.Symbol.Name + " ";

      return string.Format( "{0} {1}{2}{3}<{4} {5}>",
        typeName, categoryName, familyName, symbolName,
        e.Id.IntegerValue, e.Name );
    }

    static int ColorToInt( Color color )
    {
      return ( (int) color.Red ) << 16
        | ( (int) color.Green ) << 8
        | (int) color.Blue;
    }

    public static int ColorTransparencyToInt( 
      Color color,
      int transparency )
    {
      Debug.Assert( 0 <= transparency,
        "expected non-negative transparency" );

      Debug.Assert( 100 >= transparency,
        "expected transparency between 0 and 100" );

      uint trgb = ( (uint) transparency << 24 )
        | (uint) ColorToInt( color );

      Debug.Assert( int.MaxValue > trgb,
        "expected trgb smaller than max int" );

      return (int) trgb;
    }

    static Color IntToColor( int rgb )
    {
      return new Color(
        (byte) ( ( rgb & 0xFF0000 ) >> 16 ),
        (byte) ( ( rgb & 0xFF00 ) >> 8 ),
        (byte) ( rgb & 0xFF ) );
    }

    public static Color IntToColorTransparency( 
      int trgb, 
      out int transparency )
    {
      transparency = (int) ( ( ( (uint) trgb ) 
        & 0xFF000000 ) >> 24 );

      return IntToColor( trgb );
    }

    static string ColorString( Color color )
    {
      return color.Red.ToString( "X2" )
        + color.Green.ToString( "X2" )
        + color.Blue.ToString( "X2" );
    }

    public static string ColorTransparencyString(
      Color color,
      int transparency )
    {
      return transparency.ToString( "X2" )
        + ColorString( color );
    }

    #region GetRevitTextColorFromSystemColor
    int GetRevitTextColorFromSystemColor( 
      System.Drawing.Color color )
    {
      return ( ( (int) color.R ) * (int) Math.Pow( 2, 0 ) 
        + ( (int) color.G ) * (int) Math.Pow( 2, 8 ) 
        + ( (int) color.B ) * (int) Math.Pow( 2, 16 ) );
    }

    void RevitTextColorFromSystemColorUsageExample()
    {
      TextNoteType tnt = null;

      // Set Revit text colour from system colour

      int color = GetRevitTextColorFromSystemColor( 
        System.Drawing.Color.Wheat );

      tnt.get_Parameter( BuiltInParameter.LINE_COLOR )
        .Set( color );
    }
    #endregion // GetRevitTextColorFromSystemColor
  }
}
