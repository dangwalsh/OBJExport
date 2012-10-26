#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Application = Autodesk.Revit.ApplicationServices.Application;
#endregion // Namespaces

namespace ObjExport
{
  [Transaction( TransactionMode.ReadOnly )]
  public class Command : IExternalCommand
  {
    /// <summary>
    /// Default colour: grey.
    /// </summary>
    public static Color DefaultColor 
      = new Color( 127, 127, 127 );

    /// <summary>
    /// Default transparency: opaque.
    /// </summary>
    public static int Opaque = 0;

    void InfoMsg( string msg )
    {
      TaskDialog.Show( "OBJ Exporter", msg );
    }

    /// <summary>
    /// Define the schedule export folder.
    /// All existing files will be overwritten.
    /// </summary>
    static string _export_folder_name = null;

    /// <summary>
    /// Select an OBJ output file in the given folder.
    /// </summary>
    /// <param name="folder">Initial folder.</param>
    /// <param name="filename">Selected filename on success.</param>
    /// <returns>Return true if a file was successfully selected.</returns>
    static bool FileSelect( 
      string folder,
      out string filename )
    {
      SaveFileDialog dlg = new SaveFileDialog();
      dlg.Title = "Select OBJ Output File";
      dlg.CheckFileExists = false;
      dlg.CheckPathExists = true;
      //dlg.RestoreDirectory = true;
      dlg.InitialDirectory = folder;
      dlg.Filter = "OBJ Files (*.obj)|*.obj|All Files (*.*)|*.*";
      bool rc = ( DialogResult.OK == dlg.ShowDialog() );
      filename = dlg.FileName;
      return rc;
    }

    /// <summary>
    /// Obsolete, since a single element may contain
    /// more than one solid that really needs 
    /// exporting, e.g. the fireplace in 
    /// rac_basic_sample_project.rvt.
    /// Replaced by the ExportSolids + ExportSolid 
    /// methods.
    /// Retrieve the first non-empty solid found for 
    /// the given element. In case it is a family 
    /// instance, it may have its own non-empty solid, 
    /// in which case we use that. 
    /// Otherwise we search the symbol geometry. 
    /// If we use the symbol geometry, we might have 
    /// to keep track of the instance transform to map 
    /// it to the actual instance project location. 
    /// Instead, we ask for transformed geometry to be 
    /// returned, so the resulting solid is already in 
    /// place.
    /// </summary>
    Solid GetSolid( Element e, Options opt )
    {
      Solid solid = null;

      GeometryElement geo = e.get_Geometry( opt );

      if( null != geo )
      {
        if( e is FamilyInstance )
        {
          geo = geo.GetTransformed(
            Transform.Identity );
        }

        GeometryInstance inst = null;
        //Transform t = Transform.Identity;

        // Some columns have no solids, and we have to
        // retrieve the geometry from the symbol; 
        // others do have solids on the instance itself
        // and no contents in the instance geometry 
        // (e.g. in rst_basic_sample_project.rvt).

        foreach( GeometryObject obj in geo )
        {
          solid = obj as Solid;

          if( null != solid
            && 0 < solid.Faces.Size )
          {
            break;
          }

          inst = obj as GeometryInstance;
        }

        if( null == solid && null != inst )
        {
          geo = inst.GetSymbolGeometry();
          //t = inst.Transform;

          foreach( GeometryObject obj in geo )
          {
            solid = obj as Solid;

            if( null != solid
              && 0 < solid.Faces.Size )
            {
              break;
            }
          }
        }
      }
      return solid;
    }

    /// <summary>
    /// Export a non-empty solid.
    /// </summary>
    bool ExportSolid( 
      IJtFaceEmitter emitter,
      Document doc,
      Solid solid,
      Color color,
      int transparency )
    {
      Material m;
      Color c;
      int t;

      foreach( Face face in solid.Faces )
      {
        m = doc.GetElement(
          face.MaterialElementId ) as Material;

        c = ( null == m ) ? color : m.Color;

        t = ( null == m ) 
          ? transparency 
          : m.Transparency;

        emitter.EmitFace( face, 
          (null == c) ? DefaultColor : c,
          t );
      }
      return true;
    }

    /// <summary>
    /// Export all non-empty solids found for 
    /// the given element. Family instances may have 
    /// their own non-empty solids, in which case 
    /// those are used, otherwise the symbol geometry.
    /// The symbol geometry could keep track of the 
    /// instance transform to map it to the actual 
    /// project location. Instead, we ask for 
    /// transformed geometry to be returned, so the 
    /// resulting solids are already in place.
    /// </summary>
    int ExportSolids( 
      IJtFaceEmitter emitter,
      Element e, 
      Options opt,
      Color color,
      int transparency )
    {
      int nSolids = 0;

      GeometryElement geo = e.get_Geometry( opt );

      Solid solid;

      if( null != geo )
      {
        Document doc = e.Document;

        if( e is FamilyInstance )
        {
          geo = geo.GetTransformed(
            Transform.Identity );
        }

        GeometryInstance inst = null;
        //Transform t = Transform.Identity;

        // Some columns have no solids, and we have to
        // retrieve the geometry from the symbol; 
        // others do have solids on the instance itself
        // and no contents in the instance geometry 
        // (e.g. in rst_basic_sample_project.rvt).

        foreach( GeometryObject obj in geo )
        {
          solid = obj as Solid;

          if( null != solid
            && 0 < solid.Faces.Size
            && ExportSolid( emitter, doc, solid, 
              color, transparency ) )
          {
            ++nSolids;
          }

          inst = obj as GeometryInstance;
        }

        if( 0 == nSolids && null != inst )
        {
          geo = inst.GetSymbolGeometry();
          //t = inst.Transform;

          foreach( GeometryObject obj in geo )
          {
            solid = obj as Solid;

            if( null != solid
              && 0 < solid.Faces.Size
              && ExportSolid( emitter, doc, solid,
                color, transparency ) )
            {
              ++nSolids;
            }
          }
        }
      }
      return nSolids;
    }

    /// <summary>
    /// Export an element, i.e. all non-empty solids
    /// encountered, and return the number of elements 
    /// exported.
    /// If the element is a group, this method is 
    /// called recursively on the group members.
    /// </summary>
    int ExportElement(
      IJtFaceEmitter emitter,
      Element e,
      Options opt,
      ref int nSolids )
    {
      Group group = e as Group;

      if( null != group )
      {
        int n = 0;

        foreach( ElementId id 
          in group.GetMemberIds() )
        {
          Element e2 = e.Document.GetElement(
            id );

          n += ExportElement( emitter, e2, opt, ref nSolids );
        }
        return n;
      }

      string desc = Util.ElementDescription( e );

      Category cat = e.Category;

      if( null == cat )
      {
        Debug.Print( "Element '{0}' has no "
          + "category.", desc );

        return 0;
      }

      Material material = cat.Material;

      // Column category has no material, maybe all
      // family instances have no defualt material,
      // so we cannot simply skip them here:

      //if( null == material )
      //{
      //  Debug.Print( "Category '{0}' of element '{1}' "
      //    + "has no material.", cat.Name, desc );

      //  return 0;
      //}

      Color color = ( null == material )
        ? null
        : material.Color;

      int transparency = ( null == material )
        ? 0
        : material.Transparency;

      //Debug.Assert( null != color,
      //  "expected a valid category material colour" );

      nSolids += ExportSolids( emitter, e, opt, color, transparency );

      return 1;
    }

    void ExportElements(
      IJtFaceEmitter emitter,
      FilteredElementCollector collector,
      Options opt )
    {
      int nElements = 0;
      int nSolids = 0;

      foreach( Element e in collector )
      {
        nElements += ExportElement( 
          emitter, e, opt, ref nSolids );
      }

      int nFaces = emitter.GetFaceCount();
      int nTriangles = emitter.GetTriangleCount();
      int nVertices = emitter.GetVertexCount();

      string msg = string.Format(
        "{0} element{1} with {2} solid{3}, "
        + "{4} face{5}, {6} triangle{7} and "
        + "{8} vertice{9} exported.",
        nElements, Util.PluralSuffix( nElements ),
        nSolids, Util.PluralSuffix( nSolids ),
        nFaces, Util.PluralSuffix( nFaces ),
        nTriangles, Util.PluralSuffix( nTriangles ),
        nVertices, Util.PluralSuffix( nVertices ) );

      InfoMsg( msg );
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      // Determine elements to export

      FilteredElementCollector collector = null;

      // Access current selection

      SelElementSet set = uidoc.Selection.Elements;
      
      int n = set.Size;

      if( 0 < n )
      {
        // If any elements were preselected,
        // export those to OBJ

        ICollection<ElementId> ids = set
          .Cast<Element>()
          .Select<Element, ElementId>( e => e.Id )
          .ToArray<ElementId>();

        collector = new FilteredElementCollector( doc, ids );
      }
      else
      {
        // If nothing was preselected, export 
        // all model elements to OBJ

        collector = new FilteredElementCollector( doc );
      }

      collector.WhereElementIsNotElementType()
          .WhereElementIsViewIndependent();

      if( null == _export_folder_name )
      {
        _export_folder_name = Path.GetTempPath();
      }

      string filename = null;

      if( !FileSelect( _export_folder_name, 
        out filename ) )
      {
        return Result.Cancelled;
      }

      _export_folder_name 
        = Path.GetDirectoryName( filename );

      ObjExporter exporter = new ObjExporter();

      Options opt = app.Create.NewGeometryOptions();

      ExportElements( exporter, collector, opt );

      exporter.ExportTo( filename );

      return Result.Succeeded;
    }
  }
}
