#region Namespaces
using System;
using Autodesk.Revit.DB;
#endregion // Namespaces

namespace ObjExport
{
  // I would have liked to define the count methods 
  // as simple properties, but that cannot be done 
  // in an interface specification, unfortunately.

  interface IJtFaceEmitter
  {
    /// <summary>
    /// Emit a face with a specified colour.
    /// </summary>
    int EmitFace( 
      Face face, 
      Color color,
      int transparency );

    /// <summary>
    /// Return the final triangle count 
    /// after processing all faces.
    /// </summary>
    int GetFaceCount();

    /// <summary>
    /// Return the final triangle count 
    /// after processing all faces.
    /// </summary>
    int GetTriangleCount();

    /// <summary>
    /// Return the final vertex count 
    /// after processing all faces.
    /// </summary>
    int GetVertexCount();
  }
}
