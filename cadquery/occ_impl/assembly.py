from math import degrees
from typing import Any
from typing import Dict
from typing import Iterable
from typing import Iterator
from typing import List
from typing import Optional
from typing import Tuple
from typing import Union
from typing import cast
from typing import overload

from OCP.BOPAlgo import BOPAlgo_GlueEnum
from OCP.BOPAlgo import BOPAlgo_MakeConnected
from OCP.BRepAlgoAPI import BRepAlgoAPI_Fuse
from OCP.Quantity import Quantity_ColorRGBA
from OCP.TCollection import TCollection_AsciiString
from OCP.TCollection import TCollection_ExtendedString
from OCP.TCollection import TCollection_HAsciiString
from OCP.TDF import TDF_Label
from OCP.TDataStd import TDataStd_Name
from OCP.TDocStd import TDocStd_Document
from OCP.TopLoc import TopLoc_Location
from OCP.TopTools import TopTools_ListOfShape
from OCP.TopoDS import TopoDS_Shape
from OCP.XCAFApp import XCAFApp_Application
from OCP.XCAFDoc import XCAFDoc_ColorGen
from OCP.XCAFDoc import XCAFDoc_ColorType
from OCP.XCAFDoc import XCAFDoc_DocumentTool
from OCP.XCAFDoc import XCAFDoc_Material
from OCP.XCAFDoc import XCAFDoc_MaterialTool
from OCP.XCAFDoc import XCAFDoc_VisMaterial
from OCP.XCAFDoc import XCAFDoc_VisMaterialPBR
from typing_extensions import Protocol
from vtkmodules.vtkCommonDataModel import VTK_LINE
from vtkmodules.vtkCommonDataModel import VTK_TRIANGLE
from vtkmodules.vtkCommonDataModel import VTK_VERTEX
from vtkmodules.vtkFiltersExtraction import vtkExtractCellsByType
from vtkmodules.vtkRenderingCore import vtkActor
from vtkmodules.vtkRenderingCore import vtkPolyDataMapper as vtkMapper
from vtkmodules.vtkRenderingCore import vtkRenderer

from .exporters.vtk import toString
from .geom import Location
from .shapes import Compound
from .shapes import Shape
from .shapes import Solid
from ..cq import Workplane

# type definitions
AssemblyObjects = Union[Shape, Workplane, None]


class Color(object):
    """
    Wrapper for the OCCT color object Quantity_ColorRGBA.
    """

    wrapped: Quantity_ColorRGBA

    @overload
    def __init__(self, name: str):
        """
        Construct a Color from a name.

        :param name: name of the color, e.g. green
        """
        ...

    @overload
    def __init__(self, r: float, g: float, b: float, a: float = 0):
        """
        Construct a Color from RGB(A) values.

        :param r: red value, 0-1
        :param g: green value, 0-1
        :param b: blue value, 0-1
        :param a: alpha value, 0-1 (default: 0)
        """
        ...

    @overload
    def __init__(self):
        """
        Construct a Color with default value.
        """
        ...

    def __init__(self, *args, **kwargs):

        if len(args) == 0:
            self.wrapped = Quantity_ColorRGBA()
        elif len(args) == 1:
            self.wrapped = Quantity_ColorRGBA()
            exists = Quantity_ColorRGBA.ColorFromName_s(args[0], self.wrapped)
            if not exists:
                raise ValueError(f"Unknown color name: {args[0]}")
        elif len(args) == 3:
            r, g, b = args
            self.wrapped = Quantity_ColorRGBA(r, g, b, 1)
            if kwargs.get("a"):
                self.wrapped.SetAlpha(kwargs.get("a"))
        elif len(args) == 4:
            r, g, b, a = args
            self.wrapped = Quantity_ColorRGBA(r, g, b, a)
        else:
            raise ValueError(f"Unsupported arguments: {args}, {kwargs}")

    def toTuple(self) -> Tuple[float, float, float, float]:
        """
        Convert Color to RGB tuple.
        """
        a = self.wrapped.Alpha()
        rgb = self.wrapped.GetRGB()

        return (rgb.Red(), rgb.Green(), rgb.Blue(), a)


class Material:
    wrapped: XCAFDoc_Material
    wrapped_vis: XCAFDoc_VisMaterial

    @classmethod
    def from_ocp(cls, material: XCAFDoc_Material, vis_material: XCAFDoc_VisMaterial):
        self = cls.__new__(cls)
        self.wrapped = material
        self.wrapped_vis = vis_material
        return self

    def __init__(
            self,
            name: str,
            description: str = "",
            density: float = 0.0,
            base_color: "Color" = None,
            metalness: float = 0.0,
            roughness: float = 0.0,
            refraction_index: float = 0.0
    ):
        self.wrapped = XCAFDoc_Material()
        self.wrapped_vis = XCAFDoc_VisMaterial()
        self.wrapped_vis.SetPbrMaterial(XCAFDoc_VisMaterialPBR())

        self.name = name
        self.description = description
        self.density = density
        self.base_color = base_color or Color(0.8, 0.8, 0.8, 1.0)
        self.metalness = metalness
        self.roughness = roughness
        self.refraction_index = refraction_index

    def get_label(self, material_tool: XCAFDoc_MaterialTool):
        params = (
            self.wrapped.GetName() or TCollection_HAsciiString(""),
            self.wrapped.GetDescription() or TCollection_HAsciiString(""),
            self.wrapped.GetDensity() or 0.0,
            self.wrapped.GetDensName() or TCollection_HAsciiString(""),
            self.wrapped.GetDensValType() or TCollection_HAsciiString("")
        )
        return material_tool.AddMaterial(*params)

    @property
    def name(self) -> str:
        return self.wrapped.GetName().ToCString()

    @name.setter
    def name(self, value: str):
        self.wrapped.Set(
            TCollection_HAsciiString(value),
            self.wrapped.GetDescription(),
            self.wrapped.GetDensity(),
            self.wrapped.GetDensName(),
            self.wrapped.GetDensValType()
        )

    @property
    def description(self) -> str:
        return self.wrapped.GetDescription().ToCString()

    @description.setter
    def description(self, value: str):
        self.wrapped.Set(
            self.wrapped.GetName(),
            TCollection_HAsciiString(value),
            self.wrapped.GetDensity(),
            self.wrapped.GetDensName(),
            self.wrapped.GetDensValType()
        )

    @property
    def density(self) -> float:
        return self.wrapped.GetDensity()

    @density.setter
    def density(self, value: float):
        self.wrapped.Set(
            self.wrapped.GetName(),
            self.wrapped.GetDescription(),
            value,
            self.wrapped.GetDensName(),
            self.wrapped.GetDensValType()
        )

    @property
    def base_color(self) -> Color:
        pbr = self.wrapped_vis.PbrMaterial()
        color = Color()
        color.wrapped = pbr.BaseColor
        return color

    @base_color.setter
    def base_color(self, color: Color):
        pbr = self.wrapped_vis.PbrMaterial()
        pbr.BaseColor = color.wrapped
        self.wrapped_vis.SetPbrMaterial(pbr)

    @property
    def metalness(self) -> float:
        return self.wrapped_vis.PbrMaterial().Metallic

    @metalness.setter
    def metalness(self, value: float):
        pbr = self.wrapped_vis.PbrMaterial()
        pbr.Metallic = value
        self.wrapped_vis.SetPbrMaterial(pbr)

    @property
    def roughness(self) -> float:
        return self.wrapped_vis.PbrMaterial().Roughness

    @roughness.setter
    def roughness(self, value: float):
        pbr = self.wrapped_vis.PbrMaterial()
        pbr.Roughness = value
        self.wrapped_vis.SetPbrMaterial(pbr)

    @property
    def refraction_index(self) -> float:
        return self.wrapped_vis.PbrMaterial().RefractionIndex

    @refraction_index.setter
    def refraction_index(self, value: float):
        pbr = self.wrapped_vis.PbrMaterial()
        pbr.RefractionIndex = value
        self.wrapped_vis.SetPbrMaterial(pbr)


class AssemblyProtocol(Protocol):
    @property
    def loc(self) -> Location:
        ...

    @loc.setter
    def loc(self, value: Location) -> None:
        ...

    @property
    def name(self) -> str:
        ...

    @property
    def parent(self) -> Optional["AssemblyProtocol"]:
        ...

    @property
    def color(self) -> Optional[Color]:
        ...

    @property
    def material(self) -> Optional[Material]:
        ...

    @property
    def obj(self) -> AssemblyObjects:
        ...

    @property
    def shapes(self) -> Iterable[Shape]:
        ...

    @property
    def children(self) -> Iterable["AssemblyProtocol"]:
        ...

    def traverse(self) -> Iterable[Tuple[str, "AssemblyProtocol"]]:
        ...

    def __iter__(
            self,
            loc: Optional[Location] = None,
            name: Optional[str] = None,
            color: Optional[Color] = None,
    ) -> Iterator[Tuple[Shape, str, Location, Optional[Color]]]:
        ...


def setName(l: TDF_Label, name: str, tool):
    TDataStd_Name.Set_s(l, TCollection_ExtendedString(name))


def setColor(l: TDF_Label, color: Color, tool):
    tool.SetColor(l, color.wrapped, XCAFDoc_ColorType.XCAFDoc_ColorSurf)


def toCAF(
        assy: AssemblyProtocol,
        coloredSTEP: bool = False,
        mesh: bool = False,
        tolerance: float = 1e-3,
        angularTolerance: float = 0.1,
) -> Tuple[TDF_Label, TDocStd_Document]:
    # prepare a doc
    app = XCAFApp_Application.GetApplication_s()

    doc = TDocStd_Document(TCollection_ExtendedString("XmlOcaf"))
    app.InitDocument(doc)

    tool = XCAFDoc_DocumentTool.ShapeTool_s(doc.Main())
    tool.SetAutoNaming_s(False)
    ctool = XCAFDoc_DocumentTool.ColorTool_s(doc.Main())
    material_tool = XCAFDoc_DocumentTool.MaterialTool_s(doc.Main())
    vis_material_tool = XCAFDoc_DocumentTool.VisMaterialTool_s(doc.Main())

    # used to store labels with unique part-color combinations
    unique_objs: Dict[Tuple[Color, AssemblyObjects], TDF_Label] = {}
    # used to cache unique, possibly meshed, compounds; allows to avoid redundant meshing operations if same object is referenced multiple times in an assy
    compounds: Dict[AssemblyObjects, Compound] = {}
    materials: Dict[Material, TDF_Label] = {}
    visualization_materials: Dict[Material, TDF_Label] = {}

    def _toCAF(el, ancestor, color) -> TDF_Label:
        nonlocal ctool, material_tool, unique_objs, compounds, materials, visualization_materials

        # create a subassy
        subassy = tool.NewShape()
        setName(subassy, el.name, tool)

        # define the current color
        current_color = el.color if el.color else color

        # add a leaf with the actual part if needed
        if el.obj:
            # get/register unique parts referenced in the assy
            key0 = (current_color, el.obj)  # (color, shape)
            key1 = el.obj  # shape

            if key0 in unique_objs:
                lab = unique_objs[key0]
            else:
                lab = tool.NewShape()
                if key1 in compounds:
                    compound = compounds[key1].copy(mesh)
                else:
                    compound = Compound.makeCompound(el.shapes)
                    if mesh:
                        compound.mesh(tolerance, angularTolerance)

                    compounds[key1] = compound

                tool.SetShape(lab, compound.wrapped)
                setName(lab, f"{el.name}_part", tool)
                unique_objs[key0] = lab

                # handle colors when exporting to STEP
                if coloredSTEP and current_color:
                    setColor(lab, current_color, ctool)

                # Handle materials
                if el.material:
                    material = el.material
                    if material in materials:
                        material_label = materials[material]
                    else:
                        material_label = material.get_label(material_tool)
                        materials[material] = material_label

                    if material in visualization_materials:
                        visualization_material_label = visualization_materials[material]
                    else:
                        visualization_material_label = vis_material_tool.AddMaterial(
                            material.wrapped_vis.BackupCopy(),
                            TCollection_AsciiString(material.name)
                        )

                    material_tool.SetMaterial(lab, material_label)
                    vis_material_tool.SetShapeMaterial(lab, visualization_material_label)

            tool.AddComponent(subassy, lab, TopLoc_Location())

        # handle colors when *not* exporting to STEP
        if not coloredSTEP and current_color:
            setColor(subassy, current_color, ctool)

        # add children recursively
        for child in el.children:
            _toCAF(child, subassy, current_color)

        if ancestor:
            # add the current subassy to the higher level assy
            tool.AddComponent(ancestor, subassy, el.loc.wrapped)

        return subassy

    # process the whole assy recursively
    top = _toCAF(assy, None, None)

    tool.UpdateAssemblies()

    return top, doc


def toVTK(
    assy: AssemblyProtocol,
    color: Tuple[float, float, float, float] = (1.0, 1.0, 1.0, 1.0),
    tolerance: float = 1e-3,
    angularTolerance: float = 0.1,
) -> vtkRenderer:
    renderer = vtkRenderer()

    for shape, _, loc, col_ in assy:
        col = col_.toTuple() if col_ else color
        trans, rot = loc.toTuple()

        data = shape.toVtkPolyData(tolerance, angularTolerance)

        # extract faces
        extr = vtkExtractCellsByType()
        extr.SetInputDataObject(data)

        extr.AddCellType(VTK_LINE)
        extr.AddCellType(VTK_VERTEX)
        extr.Update()
        data_edges = extr.GetOutput()

        # extract edges
        extr = vtkExtractCellsByType()
        extr.SetInputDataObject(data)

        extr.AddCellType(VTK_TRIANGLE)
        extr.Update()
        data_faces = extr.GetOutput()

        # remove normals from edges
        data_edges.GetPointData().RemoveArray("Normals")

        # add both to the renderer
        mapper = vtkMapper()
        mapper.AddInputDataObject(data_faces)

        actor = vtkActor()
        actor.SetMapper(mapper)
        actor.SetPosition(*trans)
        actor.SetOrientation(*map(degrees, rot))
        actor.GetProperty().SetColor(*col[:3])
        actor.GetProperty().SetOpacity(col[3])

        renderer.AddActor(actor)

        mapper = vtkMapper()
        mapper.AddInputDataObject(data_edges)

        actor = vtkActor()
        actor.SetMapper(mapper)
        actor.SetPosition(*trans)
        actor.SetOrientation(*map(degrees, rot))
        actor.GetProperty().SetColor(0, 0, 0)
        actor.GetProperty().SetLineWidth(2)

        renderer.AddActor(actor)

    return renderer


def toJSON(
        assy: AssemblyProtocol,
        color: Tuple[float, float, float, float] = (1.0, 1.0, 1.0, 1.0),
        tolerance: float = 1e-3,
) -> List[Dict[str, Any]]:
    """
    Export an object to a structure suitable for converting to VTK.js JSON.
    """

    rv = []

    for shape, _, loc, col_ in assy:
        val: Any = {}

        data = toString(shape, tolerance)
        trans, rot = loc.toTuple()

        val["shape"] = data
        val["color"] = col_.toTuple() if col_ else color
        val["position"] = trans
        val["orientation"] = rot

        rv.append(val)

    return rv


def toFusedCAF(
        assy: AssemblyProtocol, glue: bool = False, tol: Optional[float] = None,
) -> Tuple[TDF_Label, TDocStd_Document]:
    """
    Converts the assembly to a fused compound and saves that within the document
    to be exported in a way that preserves the face colors. Because of the use of
    boolean operations in this method, performance may be slow in some cases.

    :param assy: Assembly that is being converted to a fused compound for the document.
    """

    # Prepare the document
    app = XCAFApp_Application.GetApplication_s()
    doc = TDocStd_Document(TCollection_ExtendedString("XmlOcaf"))
    app.InitDocument(doc)

    # Shape and color tools
    shape_tool = XCAFDoc_DocumentTool.ShapeTool_s(doc.Main())
    color_tool = XCAFDoc_DocumentTool.ColorTool_s(doc.Main())

    # To fuse the parts of the assembly together
    fuse_op = BRepAlgoAPI_Fuse()
    args = TopTools_ListOfShape()
    tools = TopTools_ListOfShape()

    # If there is only one solid, there is no reason to fuse, and it will likely cause problems anyway
    top_level_shape = None

    # Walk the entire assembly, collecting the located shapes and colors
    shapes: List[Shape] = []
    colors = []

    for shape, _, loc, color in assy:
        shapes.append(shape.moved(loc).copy())
        colors.append(color)

    # Initialize with a dummy value for mypy
    top_level_shape = cast(TopoDS_Shape, None)

    # If the tools are empty, it means we only had a single shape and do not need to fuse
    if not shapes:
        raise Exception(f"Error: Assembly {assy.name} has no shapes.")
    elif len(shapes) == 1:
        # There is only one shape and we only need to make sure it is a Compound
        # This seems to be needed to be able to add subshapes (i.e. faces) correctly
        sh = shapes[0]
        if sh.ShapeType() != "Compound":
            top_level_shape = Compound.makeCompound((sh,)).wrapped
        elif sh.ShapeType() == "Compound":
            sh = sh.fuse(glue=glue, tol=tol)
            top_level_shape = Compound.makeCompound((sh,)).wrapped
            shapes = [sh]
    else:
        # Set the shape lists up so that the fuse operation can be performed
        args.Append(shapes[0].wrapped)

        for shape in shapes[1:]:
            tools.Append(shape.wrapped)

        # Allow the caller to configure the fuzzy and glue settings
        if tol:
            fuse_op.SetFuzzyValue(tol)
        if glue:
            fuse_op.SetGlue(BOPAlgo_GlueEnum.BOPAlgo_GlueShift)

        fuse_op.SetArguments(args)
        fuse_op.SetTools(tools)
        fuse_op.Build()

        top_level_shape = fuse_op.Shape()

    # Add the fused shape as the top level object in the document
    top_level_lbl = shape_tool.AddShape(top_level_shape, False)
    TDataStd_Name.Set_s(top_level_lbl, TCollection_ExtendedString(assy.name))

    # Walk the assembly->part->shape->face hierarchy and add subshapes for all the faces
    for color, shape in zip(colors, shapes):
        for face in shape.Faces():
            # See if the face can be treated as-is
            cur_lbl = shape_tool.AddSubShape(top_level_lbl, face.wrapped)
            if color and not cur_lbl.IsNull() and not fuse_op.IsDeleted(face.wrapped):
                color_tool.SetColor(cur_lbl, color.wrapped, XCAFDoc_ColorGen)

            # Handle any modified faces
            modded_list = fuse_op.Modified(face.wrapped)

            for mod in modded_list:
                # Add the face as a subshape and set its color to match the parent assembly component
                cur_lbl = shape_tool.AddSubShape(top_level_lbl, mod)
                if color and not cur_lbl.IsNull() and not fuse_op.IsDeleted(mod):
                    color_tool.SetColor(cur_lbl, color.wrapped, XCAFDoc_ColorGen)

            # Handle any generated faces
            gen_list = fuse_op.Generated(face.wrapped)

            for gen in gen_list:
                # Add the face as a subshape and set its color to match the parent assembly component
                cur_lbl = shape_tool.AddSubShape(top_level_lbl, gen)
                if color and not cur_lbl.IsNull():
                    color_tool.SetColor(cur_lbl, color.wrapped, XCAFDoc_ColorGen)

    return top_level_lbl, doc


def imprint(assy: AssemblyProtocol) -> Tuple[Shape, Dict[Shape, Tuple[str, ...]]]:
    """
    Imprint all the solids and construct a dictionary mapping imprinted solids to names from the input assy.
    """

    # make the id map
    id_map = {}

    for obj, name, loc, _ in assy:
        for s in obj.moved(loc).Solids():
            id_map[s] = name

    # connect topologically
    bldr = BOPAlgo_MakeConnected()
    bldr.SetRunParallel(True)
    bldr.SetUseOBB(True)

    for obj in id_map:
        bldr.AddArgument(obj.wrapped)

    bldr.Perform()
    res = Shape(bldr.Shape())

    # make the connected solid -> id map
    origins: Dict[Shape, Tuple[str, ...]] = {}

    for s in res.Solids():
        ids = tuple(id_map[Solid(el)] for el in bldr.GetOrigins(s.wrapped))
        # if GetOrigins yields nothing, solid was not modified
        origins[s] = ids if ids else (id_map[s],)

    return res, origins
