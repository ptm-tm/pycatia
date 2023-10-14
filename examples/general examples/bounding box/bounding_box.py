"""
    Example - Bounding Box version 1 - 001:
    User can select axis system and any geometry
    For create bounding box with interia axis:
        user must run 
    .. code
        c:Measure Intertia
    and press button 'Create geometry' and select 'Axis System'.
    Use this new Axis system to create bounding box.
    Create Bounding box of selected geometry:
        1.Wireframe bounding box
        2.Surface bounding box
        3.Solid bounding box
    .. warning:
        Catia must be run and document open and active.
        Part document must content any geometry and one axis system
        Some critical errors send message to console
    .. note:
        Need add cylindrical bounding box
"""

##########################################################
# insert syspath to project folder so examples can be run.
# for development purposes.
import os
import sys

sys.path.insert(0, os.path.abspath('..\\pycatia'))
##########################################################
# pylint: disable=C0413
# pylint error wrong import position
import pythoncom
from pycatia.mec_mod_interfaces.axis_system import AxisSystem
from pycatia.mec_mod_interfaces.body import Body
from pycatia.mec_mod_interfaces.part import Part
from pycatia.mec_mod_interfaces.part_document import PartDocument
from pycatia.in_interfaces.reference import Reference
from pycatia.space_analyses_interfaces.spa_workbench import SPAWorkbench
from pycatia import catia
from pycatia.exception_handling.exceptions import CATIAApplicationException
from pycatia.version import version
from pycatia.knowledge_interfaces.parameter_set import Parameters
from pycatia.knowledge_interfaces.parameter_set import ParameterSet
from pycatia.knowledge_interfaces.parameter_set import ParameterSets



"""
if version <='0.6.1':
    sys.exit('version pycatia must be > 0.6.1. Current version is='+version)
"""

# TODO do not work with existing params
# TODO add name to bounding box params


__author__ = '[ptm] by plm-forum.ru'
__status__ = 'beta'

def offset_bounding_box(catia_app: catia) -> tuple():
    """
    User input offsets for bounding box
    All numbers must be float

    Returns
    -------
    tuple
        (x_max,x_min,y_max,y_min,z_max,z_min).
        where:
        x_max:  offets for max value along [x]-direction
        x_min:   offets for min value along [-x]-direction
        y_max:  offets for max value along [y]-direction
        y_min:  offets for min value along [-y]-direction
        z_max:  offets for max value along [z]-direction
        z_min:  offets for min value along [-z]-direction

    """

    error_code = True
    while error_code:
        try:
            res_x_max = float(catia_app.input_box(
                'Input offset Xmax', 'Input offsets', 20))
            res_x_min = float(catia_app.input_box(
                'Input offset Xmin', 'Input offsets', 20))
            res_y_max = float(catia_app.input_box(
                'Input offset Ymax', 'Input offsets', 20))
            res_y_min = float(catia_app.input_box(
                'Input offset Ymin', 'Input offsets', 20))
            res_z_max = float(catia_app.input_box(
                'Input offset Zmax', 'Input offsets', 20))
            res_z_min = float(catia_app.input_box(
                'Input offset Zmin', 'Input offsets', 20))
            error_code = False

        except ValueError:
            match caa.message_box('Wrong input! you must input numbers. You want'
                                  're-input offsets', 5, 'Input error!'):
                case 4:
                    error_code = True
                case 2:
                    error_code = False
                    caa.message_box('Exit')
                    sys.exit('Abort')

    return (res_x_max, res_x_min, res_y_max, res_y_min, res_z_max, res_z_min)



def axis_references(input_part: Part, input_axis: AxisSystem) -> tuple:
    """
    Return tuple of references for planes and axis

    Parameters
    ----------
    input_part : Part
        current Part() document.
    input_axis : AxisSystem
        AxisSystem in Part().

    Returns
    -------
    tuple
        (aX,aY,aZ,pXY,pXZ,pYZ)
        where:
            aX:     X-axis Reference()
            aY:     Y-axis Reference()
            aZ:     Z-axis Reference()
            pXY:    XY-plane Reference()
            pXZ:    XZ-plane Reference()
            pYZ:    YZ-Plane Reference()
    """

    s_x_axis = f'REdge:(Edge:(Face:(Brp:({input_axis.name};1);'\
        f'None:();Cf11:());Face:(Brp:({input_axis.name};3);'\
        'None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'

    s_y_axis = f'REdge:(Edge:(Face:(Brp:({input_axis.name};2);'\
        f'None:();Cf11:());Face:(Brp:({input_axis.name};1);'\
        'None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'
    s_z_axis = f'REdge:(Edge:(Face:(Brp:({input_axis.name};3);'\
        f'None:();Cf11:());Face:(Brp:({input_axis.name};2);'\
        'None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'

    s_xy_plane = f'RSur:(Face:(Brp:({input_axis.name};1);None:();Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'
    s_xz_plane = f'RSur:(Face:(Brp:({input_axis.name};3);None:();Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'
    s_yz_plane = f'RSur:(Face:(Brp:({input_axis.name};2);None:();Cf11:());'\
        'WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;'\
        'MFBRepVersion_CXR14)'

    res0 = input_part.create_reference_from_b_rep_name(s_x_axis, Axis_System)
    res1 = input_part.create_reference_from_b_rep_name(s_y_axis, Axis_System)
    res2 = input_part.create_reference_from_b_rep_name(s_z_axis, Axis_System)

    res3 = input_part.create_reference_from_b_rep_name(s_xy_plane, Axis_System)
    res4 = input_part.create_reference_from_b_rep_name(s_xz_plane, Axis_System)
    res5 = input_part.create_reference_from_b_rep_name(s_yz_plane, Axis_System)
    return (res0, res1, res2, res3, res4, res5)


def measure_between_planes(
        plane_1: Reference, plane_2: Reference, spa_wb: SPAWorkbench) -> float:
    """
        Return minimum distance between 2 planes
        All planes must bu updated!!!
    Args:
        plane_1 (Reference): first plane
        plane_2 (Reference): second plane
        spa_wb (SPAWorkbench): document SPAWorkbench

    Returns:
        float: result
    """

    mes_plane1 = spa_wb.get_measurable(plane_1)
    min_distance = mes_plane1.get_minimum_distance(plane_2)
    return min_distance


def lut(num1: float, num2: float) -> bool:
    """
    function for control user input numbers

    Parameters
    ----------
    num1 : float
        offset + dir.
    num2 : float
        offset - dir.

    Returns
    -------
    bool
        True    - Some error
        False   - all fine

    """
    return not bool(num1 + num2)


# import headers

try:
    caa = catia()
    documents = caa.documents
    document = PartDocument(caa.active_document.com_object)
except CATIAApplicationException as e:
    print(e.message)
    print('CATIA not started or document not ' +
          'opened or started several CATIA sessions')
    sys.exit(e.message)


# Input offset to bounding box

offsets = offset_bounding_box(caa)

Offset_X_min = offsets[0]
Offset_X_max = offsets[1]

Offset_Y_min = offsets[2]
Offset_Y_max = offsets[3]

Offset_Z_min = offsets[4]
Offset_Z_max = offsets[5]

message_promt = ('Creating Bounding box with offset:\n' +
                 f'Xmin={Offset_X_min}\nXmax={Offset_X_max}\n' +
                 f'Ymin={Offset_Y_min}\nXmax={Offset_Y_max}\n' +
                 f'Zmin={Offset_Z_min}\nXmax={Offset_Z_max}')

return_value = caa.message_box(message_promt, 4, title='Warring')
if return_value == 7:
    print('Canceled by user')
    sys.exit('Canceled by user')

if document.is_part:
    # need to autocomplete
    part_document = Part(document.part.com_object)
    selection = document.selection

    try:
        part_document.update()
    except CATIAApplicationException as e:
        print(e.message)
        print('Part document must be without errors!')
        sys.exit('Part document must be without errors!')

    hsf = part_document.hybrid_shape_factory

    spa = document.spa_workbench()

    caa.message_box('Select a Axis System', 0, title='Selection promt')
    sFilter = ('AxisSystem',)
    sStatus = selection.select_element2(sFilter, 'select a  local axis', True)
    if sStatus == 'Cancel':
        caa.message_box(
            'Wrong selection! Application will closed!', 16, title='Warring')
        sys.exit('Wrong selection! Application will closed!')
    Axis_System = AxisSystem(selection.item(1).value.com_object)
    selection.clear()

    ref_axis = axis_references(part_document, Axis_System)

    Hybrid_Shape_D1 = hsf.add_new_direction(ref_axis[0])
    Hybrid_Shape_D2 = hsf.add_new_direction(ref_axis[1])
    Hybrid_Shape_D3 = hsf.add_new_direction(ref_axis[2])

    ref_XY = ref_axis[3]
    ref_XZ = ref_axis[4]
    ref_YZ = ref_axis[5]

    # promt user select face
    caa.message_box('Select a HybridBodies', 0, title='Selection promt')

    #sFilter = ('Body', 'HybridShape', 'Face')
    sFilter = ('TriDim','BiDim',)
    sStatus = selection.select_element2(sFilter, 'select a HybridBody', False)
    if sStatus == 'Cancel':
        sys.exit('HybridBodies not select')

    hb = Body(selection.item(1).value.com_object)
    reference1 = part_document.create_reference_from_object(hb)
    selection.clear()

    # check plane parallel axis
    try:
        app_measure = spa.get_measurable(hb)
        pln = app_measure.get_plane()
        angle_xy = app_measure.get_angle_between(ref_XY)
        angle_xz = app_measure.get_angle_between(ref_XZ)
        angle_yz = app_measure.get_angle_between(ref_YZ)

        z_parallel = (angle_xy == 0) and lut(Offset_Z_max, Offset_Z_min)
        y_parallel = (angle_xz == 0) and lut(Offset_Y_max, Offset_Y_min)
        x_parallel = (angle_yz == 0) and lut(Offset_X_max, Offset_X_min)

        # ugly but work
        # check parallel and exit
        if z_parallel:
            STATUS = 'Surface is plane and parallel to Z_plane!\n'\
                'Offset_Y is equal 0!\n'\
                'Apps will close!'
            caa.message_box(STATUS, 16, 'Critical error!')
            sys.exit(STATUS)
        if y_parallel:
            STATUS = 'Surface is plane and parallel to Y_plane!\n'\
                'Offset_Y is equal 0!\n'\
                'Apps will close!'
            caa.message_box(STATUS, 16, 'Critical error!')
            sys.exit(STATUS)
        if x_parallel:
            STATUS = 'Surface is plane and parallel to X_plane!\n'\
                'Offset_X is equal 0!\n'\
                'Apps will close!'
            caa.message_box(STATUS, 16, 'Critical error!')
            sys.exit(STATUS)
        # pylint: disable=E1101
    except pythoncom.com_error:
        #pylint error but need to detect not a plane
        pass

    #Add dimension
    # Read number of bounding box
    try:
        product=document.product
        user_prop=product.user_ref_properties
        j=int(user_prop.item('j').value)+1
        user_prop.item('j').value=j
    except CATIAApplicationException as e:
        print(e.message)
        print('will use j=0')
        j=0
        user_prop.create_integer('j', 0)

    # Create structure for geometry
    #   |-Bounding_box.X            :solid Body
    #   |-GSD Bounding Box.X        :geometrical sets main
    #   |---|-Extreme Points.X
    #   |---|-Planes.X
    #   |---|-Base Lines.X
    #   |---|-Points.X
    #   |---|-Edge.X
    #   |---|-Surfaces.X
    #   |-profile_pad.X             :output profile for solid Body
    #   |-Wireframe_Bounding_Box.X  :output edge
    #   |-Surface_Bounding_box.X    :output surfase

    oBodies = part_document.bodies
    j = oBodies.count
    body1 = oBodies.add()
    body1.name = f'Bounding_Box.{j}'

    hybridBodies1 = part_document.hybrid_bodies
    hybridBody_main = hybridBodies1.add()
    hybridBody_main.name = f'GSD Bounding Box.{j}'
    part_document.in_work_object = hybridBody_main.hybrid_bodies
    hybridBody_Extreme_Points = hybridBody_main.hybrid_bodies.add()
    hybridBody_Extreme_Points.name = f'Extreme Points.{j}'
    hybridBody_Planes = hybridBody_main.hybrid_bodies.add()
    hybridBody_Planes.name = f'Planes.{j}'
    hybridBody_Base_Lines = hybridBody_main.hybrid_bodies.add()
    hybridBody_Base_Lines.name = f'Base Lines.{j}'
    hybridBody_Points = hybridBody_main.hybrid_bodies.add()
    hybridBody_Points.name = f'Points.{j}'
    hybridBody_Edge = hybridBody_main.hybrid_bodies.add()
    hybridBody_Edge.name = f'Edge.{j}'
    hybridBody_Surfaces = hybridBody_main.hybrid_bodies.add()
    hybridBody_Surfaces.name = f'Surfaces.{j}'

    selection.add(hybridBody_Extreme_Points)
    selection.add(hybridBody_Planes)
    selection.add(hybridBody_Base_Lines)
    selection.add(hybridBody_Points)
    selection.add(hybridBody_Edge)
    selection.add(hybridBody_Surfaces)
    selection.vis_properties.set_show(1)
    selection.clear()

    # create 6 extremum points

    HybridShapeExtremum1 = hsf.add_new_extremum(reference1, Hybrid_Shape_D1, 1)
    HybridShapeExtremum1.direction = Hybrid_Shape_D1
    HybridShapeExtremum1.direction2 = Hybrid_Shape_D2
    HybridShapeExtremum1.direction3 = Hybrid_Shape_D3
    HybridShapeExtremum1.extremum_type = 1
    HybridShapeExtremum1.extremum_type2 = 1
    HybridShapeExtremum1.extremum_type3 = 1
    HybridShapeExtremum1.name = f'max_X.{j}'

    HybridShapeExtremum2 = hsf.add_new_extremum(reference1, Hybrid_Shape_D1, 0)
    HybridShapeExtremum2.direction = Hybrid_Shape_D1
    HybridShapeExtremum2.direction2 = Hybrid_Shape_D2
    HybridShapeExtremum2.direction3 = Hybrid_Shape_D3
    HybridShapeExtremum2.extremum_type = 0
    HybridShapeExtremum2.extremum_type2 = 0
    HybridShapeExtremum2.extremum_type3 = 0
    HybridShapeExtremum2.name = f'min_X.{j}'

    HybridShapeExtremum3 = hsf.add_new_extremum(reference1, Hybrid_Shape_D2, 1)
    HybridShapeExtremum3.direction = Hybrid_Shape_D2
    HybridShapeExtremum3.direction2 = Hybrid_Shape_D1
    HybridShapeExtremum3.direction3 = Hybrid_Shape_D3
    HybridShapeExtremum3.extremum_type = 1
    HybridShapeExtremum3.extremum_type2 = 1
    HybridShapeExtremum3.extremum_type3 = 1
    HybridShapeExtremum3.name = f'max_Y.{j}'

    HybridShapeExtremum4 = hsf.add_new_extremum(reference1, Hybrid_Shape_D2, 0)
    HybridShapeExtremum4.direction = Hybrid_Shape_D2
    HybridShapeExtremum4.direction2 = Hybrid_Shape_D1
    HybridShapeExtremum4.direction3 = Hybrid_Shape_D3
    HybridShapeExtremum4.extremum_type = 0
    HybridShapeExtremum4.extremum_type2 = 0
    HybridShapeExtremum4.extremum_type3 = 0
    HybridShapeExtremum4.name = f'min_Y.{j}'

    HybridShapeExtremum5 = hsf.add_new_extremum(reference1, Hybrid_Shape_D3, 1)
    HybridShapeExtremum5.direction = Hybrid_Shape_D3
    HybridShapeExtremum5.direction2 = Hybrid_Shape_D1
    HybridShapeExtremum5.direction3 = Hybrid_Shape_D2
    HybridShapeExtremum5.extremum_type = 1
    HybridShapeExtremum5.extremum_type2 = 1
    HybridShapeExtremum5.extremum_type3 = 1
    HybridShapeExtremum5.name = f'max_Z.{j}'

    HybridShapeExtremum6 = hsf.add_new_extremum(reference1, Hybrid_Shape_D3, 0)
    HybridShapeExtremum6.direction = Hybrid_Shape_D3
    HybridShapeExtremum6.direction2 = Hybrid_Shape_D1
    HybridShapeExtremum6.direction3 = Hybrid_Shape_D2
    HybridShapeExtremum6.extremum_type = 0
    HybridShapeExtremum6.extremum_type2 = 0
    HybridShapeExtremum6.extremum_type3 = 0
    HybridShapeExtremum6.name = f'min_Z.{j}'

    # append points to Geometrical Set

    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum1)
    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum2)
    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum3)
    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum4)
    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum5)
    hybridBody_Extreme_Points.append_hybrid_shape(HybridShapeExtremum6)

    # create 12 planes
    # 6 extremum planes

    Plane_Xmax = hsf.add_new_plane_offset_pt(ref_YZ, HybridShapeExtremum1)
    Plane_Xmax.name = f'Plane_X_max.{j}'
    Plane_Xmin = hsf.add_new_plane_offset_pt(ref_YZ, HybridShapeExtremum2)
    Plane_Xmin.name = f'Plane_X_min.{j}'

    Plane_Ymax = hsf.add_new_plane_offset_pt(ref_XZ, HybridShapeExtremum3)
    Plane_Ymax.name = f'Plane_Y_max.{j}'
    Plane_Ymin = hsf.add_new_plane_offset_pt(ref_XZ, HybridShapeExtremum4)
    Plane_Ymin.name = f'Plane_Y_min.{j}'

    Plane_Zmax = hsf.add_new_plane_offset_pt(ref_XY, HybridShapeExtremum5)
    Plane_Zmax.name = f'Plane_Z_max.{j}'
    Plane_Zmin = hsf.add_new_plane_offset_pt(ref_XY, HybridShapeExtremum6)
    Plane_Zmin.name = f'Plane_Z_min.{j}'

    hybridBody_Planes.append_hybrid_shape(Plane_Xmax)
    hybridBody_Planes.append_hybrid_shape(Plane_Xmin)
    hybridBody_Planes.append_hybrid_shape(Plane_Ymax)
    hybridBody_Planes.append_hybrid_shape(Plane_Ymin)
    hybridBody_Planes.append_hybrid_shape(Plane_Zmax)
    hybridBody_Planes.append_hybrid_shape(Plane_Zmin)

    # and 6 offset planes

    Plane_Xmax_offset = hsf.add_new_plane_offset(
        Plane_Xmax, Offset_X_max, False)
    Plane_Xmax_offset.name = f'Plane_X_max_offset.{j}'
    Plane_Xmin_offset = hsf.add_new_plane_offset(
        Plane_Xmin, Offset_X_min, True)
    Plane_Xmin_offset.name = f'Plane_X_min_offset.{j}'

    Plane_Ymax_offset = hsf.add_new_plane_offset(
        Plane_Ymax, Offset_Y_max, False)
    Plane_Ymax_offset.name = f'Plane_Y_max_offset.{j}'
    Plane_Ymin_offset = hsf.add_new_plane_offset(
        Plane_Ymin, Offset_Y_min, True)
    Plane_Ymin_offset.name = f'Plane_Y_min_offset.{j}'

    Plane_Zmax_offset = hsf.add_new_plane_offset(
        Plane_Zmax, Offset_Z_max, False)
    Plane_Zmax_offset.name = f'Plane_Z_max_offset.{j}'
    Plane_Zmin_offset = hsf.add_new_plane_offset(
        Plane_Zmin, Offset_Z_min, True)
    Plane_Zmin_offset.name = f'Plane_Z_min_offset.{j}'

    hybridBody_Planes.append_hybrid_shape(Plane_Xmax_offset)
    hybridBody_Planes.append_hybrid_shape(Plane_Xmin_offset)
    hybridBody_Planes.append_hybrid_shape(Plane_Ymax_offset)
    hybridBody_Planes.append_hybrid_shape(Plane_Ymin_offset)
    hybridBody_Planes.append_hybrid_shape(Plane_Zmax_offset)
    hybridBody_Planes.append_hybrid_shape(Plane_Zmin_offset)

    #prod_params=my_product.user
    part_param_set=part_document.parameters
    root_param_set=part_param_set.root_parameter_set

    s_param=root_param_set.s

    #TODO add name to params
    #part_param_set.name=f'Parameters of bounding box.{j}'
    part_param_set.create_dimension(f'Offset_X_max.{j}', 'LENGTH',Offset_X_max)
    part_param_set.create_dimension(f'Offset_X_min.{j}', 'LENGTH',Offset_X_min)
    part_param_set.create_dimension(f'Offset_Y_max.{j}', 'LENGTH',Offset_Y_max)
    part_param_set.create_dimension(f'Offset_Y_min.{j}', 'LENGTH',Offset_Y_min)
    part_param_set.create_dimension(f'Offset_Z_max.{j}', 'LENGTH',Offset_Z_max)
    part_param_set.create_dimension(f'Offset_Z_min.{j}', 'LENGTH',Offset_Z_min)


    #add relation to offset planes

    part_relation = part_document.relations
    part_relation.create_formula(
        f'Offset_X_max.{j}', 'Offset to X_max direction',
        Plane_Xmax_offset.offset, f'Offset_X_max.{j}')
    part_relation.create_formula(
        f'Offset_X_min.{j}', 'Offset to X_min direction',
        Plane_Xmin_offset.offset, f'Offset_X_min.{j}')
    part_relation.create_formula(
        f'Offset_Y_max.{j}', 'Offset to Y_max direction',
        Plane_Ymax_offset.offset, f'Offset_Y_max.{j}')
    part_relation.create_formula(
        f'Offset_Y_min.{j}', 'Offset to Y_min direction',
        Plane_Ymin_offset.offset, f'Offset_Y_min.{j}')
    part_relation.create_formula(
        f'Offset_Z_max.{j}', 'Offset to Z_max direction',
        Plane_Zmax_offset.offset, f'Offset_Z_max.{j}')
    part_relation.create_formula(
        f'Offset_Z_min.{j}', 'Offset to Z_min direction',
        Plane_Zmin_offset.offset, f'Offset_Z_min.{j}')


    # get bounding box measure

    # need update planes to measure
    part_document.update_object(hybridBody_Planes)

    x_length = measure_between_planes(
        part_document.create_reference_from_object(Plane_Xmax),
        part_document.create_reference_from_object(Plane_Xmin),
        spa)
    y_length = measure_between_planes(
        part_document.create_reference_from_object(Plane_Ymax),
        part_document.create_reference_from_object(Plane_Ymin),
        spa)
    z_length = measure_between_planes(
            part_document.create_reference_from_object(Plane_Zmax),
        part_document.create_reference_from_object(Plane_Zmin),
        spa)

    hybridBody_main.name = (
        f'GSD Bounding Box.{j} [{x_length:.2f}x{y_length:.2f}х{z_length:.2f}]')

    # create intersections offset planes

    # Zmin->Xmax,Xmin,Ymax,Ymin. its enough for bounding box
    # for example
    # fill= hsf.add_new_fill()
    # fill.add_bound(line1,line2,line3,line4)
    # bounding box = pad (fill) up to Zmax plane
    #
    #  Point_H0V1         Line_H1
    #            *-------------------------------*  Point_H1V1
    #            |                               |
    #            |                               |
    #            |                               |
    #            |                               |
    #            | <----- Line_V0                | <---Line_V1
    #            |                               |
    #            |                               |
    #            |                               |
    #            *-------------------------------*
    #  Point_H0V0         Line_H0                   Point_H1V0
    #
    #   Y
    #   ^
    #   |
    #   |
    #   |
    #   |
    #   0-------------> X

    Line_V1 = hsf.add_new_intersection(Plane_Xmax_offset, Plane_Zmin_offset)
    Line_V0 = hsf.add_new_intersection(Plane_Xmin_offset, Plane_Zmin_offset)
    Line_H1 = hsf.add_new_intersection(Plane_Ymax_offset, Plane_Zmin_offset)
    Line_H0 = hsf.add_new_intersection(Plane_Ymin_offset, Plane_Zmin_offset)
    Line_V1.name = f'Line_V1.{j}'
    Line_V0.name = f'Line_V0.{j}'
    Line_H1.name = f'Line_H1.{j}'
    Line_H0.name = f'Line_H0.{j}'

    # but u can use translate
    # do not intersection directly
    # its not working with cylindrycal object
    # Line_Xmax_Zmax_offset=hsf.add_new_translate(Line_Xmax_Zmin_offset,)
    # Line_Xmin_Zmax_offset=hsf.add_new_translate(Line_Xmin_Zmin_offset,)
    # Line_Ymax_Zmax_offset=hsf.add_new_translate(Line_Ymax_Zmin_offset,)
    # Line_Ymin_Zmax_offset=hsf.add_new_translate(Line_Ymin_Zmin_offset,)

    # append all in wireframe
    hybridBody_Base_Lines.append_hybrid_shape(Line_V1)
    hybridBody_Base_Lines.append_hybrid_shape(Line_V0)
    hybridBody_Base_Lines.append_hybrid_shape(Line_H1)
    hybridBody_Base_Lines.append_hybrid_shape(Line_H0)

    Point_H1V1 = hsf.add_new_intersection(Line_V1, Line_H1)
    Point_H0V1 = hsf.add_new_intersection(Line_V1, Line_H0)
    Point_H1V0 = hsf.add_new_intersection(Line_V0, Line_H1)
    Point_H0V0 = hsf.add_new_intersection(Line_V0, Line_H0)

    Point_H1V1.name = f'Point_H1V1.{j}'
    Point_H0V1.name = f'Point_H0V1.{j}'
    Point_H1V0.name = f'Point_H1V0.{j}'
    Point_H0V0.name = f'Point_H0V0.{j}'

    hybridBody_Points.append_hybrid_shape(Point_H1V1)
    hybridBody_Points.append_hybrid_shape(Point_H0V1)
    hybridBody_Points.append_hybrid_shape(Point_H1V0)
    hybridBody_Points.append_hybrid_shape(Point_H0V0)

    Point_H1V1_max = hsf.add_new_project(Point_H1V1, Plane_Zmax_offset)
    Point_H1V1_max.direction = Hybrid_Shape_D3
    Point_H0V1_max = hsf.add_new_project(Point_H0V1, Plane_Zmax_offset)
    Point_H0V1_max.direction = Hybrid_Shape_D3
    Point_H1V0_max = hsf.add_new_project(Point_H1V0, Plane_Zmax_offset)
    Point_H1V0_max.direction = Hybrid_Shape_D3
    Point_H0V0_max = hsf.add_new_project(Point_H0V0, Plane_Zmax_offset)
    Point_H0V0_max.direction = Hybrid_Shape_D3

    Point_H1V1_max.name = f'Point_H1V1_max.{j}'
    Point_H0V1_max.name = f'Point_H0V1_max.{j}'
    Point_H1V0_max.name = f'Point_H1V0_max.{j}'
    Point_H0V0_max.name = f'Point_H0V0_max.{j}'

    # For visual properties
    Point_tuple = (Point_H1V1,
                   Point_H0V1,
                   Point_H1V0,
                   Point_H0V0,
                   Point_H1V1_max,
                   Point_H0V1_max,
                   Point_H1V0_max,
                   Point_H0V0_max)

    hybridBody_Points.append_hybrid_shape(Point_H1V1_max)
    hybridBody_Points.append_hybrid_shape(Point_H0V1_max)
    hybridBody_Points.append_hybrid_shape(Point_H1V0_max)
    hybridBody_Points.append_hybrid_shape(Point_H0V0_max)

   # Create 8 lines for bounding box
   # To Line H0V0->H0V1->H1V1->H1V0->H0V0

    Line_H0V0_H0V1 = hsf.add_new_line_pt_pt(Point_H0V0, Point_H0V1)
    Line_H0V1_H1V1 = hsf.add_new_line_pt_pt(Point_H0V1, Point_H1V1)
    Line_H1V1_H1V0 = hsf.add_new_line_pt_pt(Point_H1V1, Point_H1V0)
    Line_H1V0_H0V0 = hsf.add_new_line_pt_pt(Point_H1V0, Point_H0V0)

    Line_H0V0_H0V1.name = f'Line_H0V0_H0V1.{j}'
    Line_H0V1_H1V1.name = f'Line_H0V1_H1V1.{j}'
    Line_H1V1_H1V0.name = f'Line_H1V1_H1V0.{j}'
    Line_H1V0_H0V0.name = f'Line_H1V0_H0V0.{j}'

    Line_H0V0_H0V1_Zmax = hsf.add_new_line_pt_pt(
        Point_H0V0_max, Point_H0V1_max)
    Line_H0V1_H1V1_Zmax = hsf.add_new_line_pt_pt(
        Point_H0V1_max, Point_H1V1_max)
    Line_H1V1_H1V0_Zmax = hsf.add_new_line_pt_pt(
        Point_H1V1_max, Point_H1V0_max)
    Line_H1V0_H0V0_Zmax = hsf.add_new_line_pt_pt(
        Point_H1V0_max, Point_H0V0_max)

    Line_H0V0_H0V1_Zmax.name = f'Line_H0V0_H0V1_Zmax.{j}'
    Line_H0V1_H1V1_Zmax.name = f'Line_H0V1_H1V1_Zmax.{j}'
    Line_H1V1_H1V0_Zmax.name = f'Line_H1V1_H1V0_Zmax.{j}'
    Line_H1V0_H0V0_Zmax.name = f'Line_H1V0_H0V0_Zmax.{j}'

    Line_H1V1_H1V1_max = hsf.add_new_line_pt_pt(Point_H1V1, Point_H1V1_max)
    Line_H0V1_H0V1_max = hsf.add_new_line_pt_pt(Point_H0V1, Point_H0V1_max)
    Line_H1V0_H1V0_max = hsf.add_new_line_pt_pt(Point_H1V0, Point_H1V0_max)
    Line_H0V0_H0V0_max = hsf.add_new_line_pt_pt(Point_H0V0, Point_H0V0_max)

    Line_H1V1_H1V1_max.name = f'Line_H1V1_H1V1_max.{j}'
    Line_H0V1_H0V1_max.name = f'Line_H0V1_H0V1_max.{j}'
    Line_H1V0_H1V0_max.name = f'Line_H1V0_H1V0_max.{j}'
    Line_H0V0_H0V0_max.name = f'Line_H0V0_H0V0_max.{j}'

    hybridBody_Edge.append_hybrid_shape(Line_H0V0_H0V1)
    hybridBody_Edge.append_hybrid_shape(Line_H0V1_H1V1)
    hybridBody_Edge.append_hybrid_shape(Line_H1V1_H1V0)
    hybridBody_Edge.append_hybrid_shape(Line_H1V0_H0V0)

    hybridBody_Edge.append_hybrid_shape(Line_H0V0_H0V1_Zmax)
    hybridBody_Edge.append_hybrid_shape(Line_H0V1_H1V1_Zmax)
    hybridBody_Edge.append_hybrid_shape(Line_H1V1_H1V0_Zmax)
    hybridBody_Edge.append_hybrid_shape(Line_H1V0_H0V0_Zmax)

    hybridBody_Edge.append_hybrid_shape(Line_H1V1_H1V1_max)
    hybridBody_Edge.append_hybrid_shape(Line_H0V1_H0V1_max)
    hybridBody_Edge.append_hybrid_shape(Line_H1V0_H1V0_max)
    hybridBody_Edge.append_hybrid_shape(Line_H0V0_H0V0_max)

    # Profile for pad
    profile_pad = hsf.add_new_join(Line_H0V0_H0V1, Line_H0V1_H1V1)
    profile_pad.add_element(Line_H1V1_H1V0)
    profile_pad.add_element(Line_H1V0_H0V0)
    profile_pad.set_manifold(True)
    profile_pad.set_connex(True)
    profile_pad.name = f'profile_pad.{j}'
    hybridBody_main.append_hybrid_shape(profile_pad)

    Wireframe_Bounding_Box = hsf.add_new_join(Line_H0V0_H0V1, Line_H0V1_H1V1)
    Wireframe_Bounding_Box.add_element(Line_H1V1_H1V0)
    Wireframe_Bounding_Box.add_element(Line_H1V0_H0V0)
    Wireframe_Bounding_Box.add_element(Line_H0V0_H0V1_Zmax)
    Wireframe_Bounding_Box.add_element(Line_H0V1_H1V1_Zmax)
    Wireframe_Bounding_Box.add_element(Line_H1V1_H1V0_Zmax)
    Wireframe_Bounding_Box.add_element(Line_H1V0_H0V0_Zmax)
    Wireframe_Bounding_Box.add_element(Line_H1V1_H1V1_max)
    Wireframe_Bounding_Box.add_element(Line_H0V1_H0V1_max)
    Wireframe_Bounding_Box.add_element(Line_H1V0_H1V0_max)
    Wireframe_Bounding_Box.add_element(Line_H0V0_H0V0_max)

    Wireframe_Bounding_Box.name = f'Wireframe_Bounding_Box.{j} '\
        '[{x_length:.2f}x{y_length:.2f}х{z_length:.2f}]'

    # non mainfold
    Wireframe_Bounding_Box.set_manifold(False)

    hybridBody_main.append_hybrid_shape(Wireframe_Bounding_Box)

    # Surface Bounding box
    Fill_Zmin = hsf.add_new_fill()
    Fill_Zmin.add_bound(Line_H0V0_H0V1)
    Fill_Zmin.add_bound(Line_H0V1_H1V1)
    Fill_Zmin.add_bound(Line_H1V1_H1V0)
    Fill_Zmin.add_bound(Line_H1V0_H0V0)
    Fill_Zmin.name=f'Fill_Zmin.{j}'
    hybridBody_Surfaces.append_hybrid_shape(Fill_Zmin)

    Fill_Zmax = hsf.add_new_fill()
    Fill_Zmax.add_bound(Line_H0V0_H0V1_Zmax)
    Fill_Zmax.add_bound(Line_H0V1_H1V1_Zmax)
    Fill_Zmax.add_bound(Line_H1V1_H1V0_Zmax)
    Fill_Zmax.add_bound(Line_H1V0_H0V0_Zmax)
    Fill_Zmax.name=f'Fill_Zmax.{j}'
    hybridBody_Surfaces.append_hybrid_shape(Fill_Zmax)


    # extrude dont work in 0.6.1 ver
    if version > '0.6.1':
        wall = hsf.add_new_extrude(profile_pad, 20, 0, Hybrid_Shape_D3)
        wall.first_limit_type = 2
        wall.second_limit_type = 2
        wall.compute()
        wall.first_upto_element = part_document.create_reference_from_object(
            Plane_Zmin_offset)
        wall.second_upto_element = part_document.create_reference_from_object(
            Plane_Zmax_offset)
    else:
        wall = hsf.add_new_sweep_line(profile_pad)
        wall.mode = 6
        wall.solution_no = 0
        wall.smooth_activity = False
        wall.guide_deviation_activity = False
        wall.draft_computation_mode = 0
        wall.draft_direction = Hybrid_Shape_D3
        wall_lim1_ref = part_document.create_reference_from_object(
            Plane_Zmax_offset)
        wall_lim2_ref = part_document.create_reference_from_object(
            Plane_Zmin_offset)
        wall.set_first_length_definition_type(3, wall_lim1_ref)
        wall.set_second_length_definition_type(3, wall_lim2_ref)

    wall.name = f'Wall_e..{j}'
    hybridBody_Surfaces.append_hybrid_shape(wall)

    Surface_Bounding_box = hsf.add_new_join(Fill_Zmin, wall)
    Surface_Bounding_box.add_element(Fill_Zmax)
    Surface_Bounding_box.set_manifold(True)
    Surface_Bounding_box.set_connex(True)
    Surface_Bounding_box.name = f'Surface_Bounding_box.{j} '\
        '[{x_length:.2f}x{y_length:.2f}х{z_length:.2f}]'

    hybridBody_main.append_hybrid_shape(Surface_Bounding_box)

    # solid
    # update profile to selection
    part_document.update_object(profile_pad)

    sf = part_document.shape_factory
    part_document.in_work_object = body1
    ref = part_document.create_reference_from_object(profile_pad)
    pad = sf.add_new_pad_from_ref(ref, 50)
    pad.set_direction(ref_axis[2])

    pad_1_limit_1 = pad.first_limit
    pad_1_limit_1.limit_mode = 3
    pad_1_limit_1.limiting_element = Plane_Zmax_offset
    pad_1_limit_2 = pad.second_limit
    pad_1_limit_2.limit_mode = 3
    pad_1_limit_2.limiting_element = Plane_Zmin_offset

    for pt in Point_tuple:
        selection.add(pt)
    # add vis property to point
    # part must be updated
    part_document.update_object(body1)
    part_document.update_object(hybridBody_main)

    selection.vis_properties.set_show(0)
    selection.vis_properties.set_symbol_type(12)
    selection.vis_properties.set_real_color(0, 255, 0, 0)
    selection.clear()
    selection.add(hybridBody_Points)
    selection.vis_properties.set_show(0)
    selection.clear()

    # add vis property to wireframe bounding box
    selection.add(Wireframe_Bounding_Box)
    selection.vis_properties.set_real_width(2, 0)
    selection.vis_properties.set_real_color(255, 255, 8, 0)
    selection.vis_properties.set_real_opacity(180, 0)
    selection.clear()

    # add vis property to Bounding_box Body
    selection.add(Surface_Bounding_box)
    selection.vis_properties.set_real_color(255, 128, 255, 0)
    selection.vis_properties.set_real_opacity(55, 0)
    selection.clear()

    # add vis property to surface bounding box
    part_document.in_work_object = body1
    selection.add(pad)
    selection.vis_properties.set_real_color(128, 0, 255, 0)
    selection.vis_properties.set_real_opacity(55, 0)
    selection.clear()
    body1.name=body1.name+f' [{x_length:.2f}x{y_length:.2f}х{z_length:.2f}]'

    sys.exit(f'Bounding box.{j} [{x_length:.2f}x{y_length:.2f}х{z_length:.2f}]'
             'created sucsesfull')
else:
    print('must be a part')
    caa.message_box('Wrong selection! Application will closed!',
                    16, title='Warring')
    sys.exit('must be a part')
