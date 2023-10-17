# -*- coding: utf-8 -*-
"""
    Example - add parameters and parameters sets:
    .. warning:
        Catia must be run and document open and active.
        Part document must content any geometry and one axis system
        Some critical errors send message to console
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
from pycatia.mec_mod_interfaces.part import Part

from pycatia.mec_mod_interfaces.part_document import PartDocument
from pycatia import catia
from pycatia.exception_handling.exceptions import CATIAApplicationException

__author__ = '[ptm] by plm-forum.ru'
__status__ = 'beta'

try:
    caa = catia()
    documents = caa.documents
    document = PartDocument(caa.active_document.com_object)
except CATIAApplicationException as e:
    print(e.message)
    print('CATIA not started or document not ' +
          'opened or started several CATIA sessions')
    caa.application.visible=True
    sys.exit(e.message)

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

#prod_params=my_product.user

"""
    Dim parameters1 As Parameters
    Set parameters1 = part1.Part.Parameters
    Dim ParameterSet1 As ParameterSet
    Set ParameterSet1 = parameters1.RootParameterSet
    Dim parameterSets1 As ParameterSets
    Set parameterSets1 = parameterSet1.ParameterSets
"""
#part_param_set.create_set_of_parameters(part_param_set.root_parameter_set)

part_main_param=part_document.parameters
#get or create main param set!
main_set=part_main_param.root_parameter_set


bool_param=part_main_param.create_boolean('Bool', True)
dim_length=part_main_param.create_dimension('Dim_length', 'Length', '100')
dim_angle=part_main_param.create_dimension('Dim_angle', 'Angle', '5')
dim_integer=part_main_param.create_integer('Integer', 0)
list_param=part_main_param.create_list('List')

list_param.value_list.add(bool_param)
list_param.value_list.add(dim_length)
list_param.value_list.add(dim_angle)
list_param.value_list.add(dim_integer)

part_main_param.create_real('Real', 1.02)    

param_set=part_main_param.root_parameter_set
#create parameters.n with default names
sets_new=part_main_param.create_set_of_parameters(param_set)

string_param=part_main_param.create_string('String', 'testing string...')

#create named parameters set

param_sets=param_set.parameter_sets
new_set=param_sets.create_set('test')
sub_param=new_set.direct_parameters

bool_param2=sub_param.create_boolean('sub_bool', False)
bool_param2.hidden=True
sub_param.create_dimension('sub_length', 'length', 300.1)
sub_param.create_dimension('sub_angle', 'angle', 15)

print(sub_param.has_parameters())
sub_param.root_parameter_set

#st=new_set.direct_parameters
print('1')
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        