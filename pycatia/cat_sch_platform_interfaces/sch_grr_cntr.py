#! usr/bin/python3.9
"""
    Module initially auto generated using V5Automation files from CATIA V5 R28 on 2020-09-25 14:34:21.593357

    .. warning::
        The notes denoted "CAA V5 Visual Basic Help" are to be used as reference only.
        They are there as a guide as to how the visual basic / catscript functions work
        and thus help debugging in pycatia.
        
"""
from pycatia.cat_sch_platform_interfaces.sch_grr import SchGRR
from pycatia.system_interfaces.any_object import AnyObject


class SchGRRCntr(AnyObject):
    """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357)

                | System.IUnknown
                |     System.IDispatch
                |         System.CATBaseUnknown
                |             System.CATBaseDispatch
                |                 System.AnyObject
                |                     SchGRRCntr
                | 
                | Manage the graphical representation of a schematic connector.
    
    """

    def __init__(self, com_object):
        super().__init__(com_object)
        self.sch_grr_cntr = com_object

    def get_symbol(self, o_grr: SchGRR, o_e_symbol_type: int) -> None:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357))
                | o Sub GetSymbol(SchGRR oGRR,
                | CatSchIDLCntrSymbolType oESymbolType)
                | 
                |     Get the graphical primitive of a connector.
                | 
                |     Parameters:
                | 
                |         oGRR
                |             The graphical primitive (ditto) used to represent a connector.
                |             
                |         oESymbolType
                |             Connector symbol type such as: point, point/vector, OnOffSheet,
                |             LineBoundary. 
                | 
                |     Example:
                |
                |          Dim objThisIntf As SchGRRCntr
                |          Dim objArg1 As SchGRR
                | 
                |           ...
                |          objThisIntf.GetSymbolobjArg1,CatSchIDLCntrSymbolType_Enum

        :param SchGRR o_grr:
        :param int o_e_symbol_type: enum cat_sch_idl_cntr_symbol_type
        :rtype: None
        """
        return self.sch_grr_cntr.GetSymbol(o_grr.com_object, o_e_symbol_type)
        # # # # Autogenerated comment: 
        # # some methods require a system service call as the methods expects a vb array object
        # # passed to it and there is no way to do this directly with python. In those cases the following code
        # # should be uncommented and edited accordingly. Otherwise completely remove all this.
        # # vba_function_name = 'get_symbol'
        # # vba_code = """
        # # Public Function get_symbol(sch_grr_cntr)
        # #     Dim oGRR (2)
        # #     sch_grr_cntr.GetSymbol oGRR
        # #     get_symbol = oGRR
        # # End Function
        # # """

        # # system_service = SystemService(self.application.SystemService)
        # # return system_service.evaluate(vba_code, 0, vba_function_name, [self.com_object])

    def remove_symbol(self) -> None:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357))
                | o Sub RemoveSymbol()
                | 
                |     Remove the graphical primitive used as the connector's symbol. The default
                |     connector's symbol type will be set to point.
                | 
                |     Example:
                | 
                |           
                | 
                |          Dim objThisIntf As SchGRRCntr
                |           ...
                |          objThisIntf.RemoveSymbol

        :rtype: None
        """
        return self.sch_grr_cntr.RemoveSymbol()

    def set_symbol(self, i_grr_symbol: SchGRR, i_e_symbol_type: int) -> None:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357))
                | o Sub SetSymbol(SchGRR iGRRSymbol,
                | CatSchIDLCntrSymbolType iESymbolType)
                | 
                |     Set the symbol or graphics used to represent a connector.
                | 
                |     Parameters:
                | 
                |         iGRRSymbol
                |             The graphical primitive (detail) to be used as the connector's
                |             symbol. iGRRSymbol can be NULL if iESymbolType is a point or point/vector.
                |             
                |         iESymbolType
                |             Connector symbol type such as: point, point/vector, OnOffSheet,
                |             LineBoundary. 
                | 
                |     Example:
                |
                |          Dim objThisIntf As SchGRRCntr
                |          Dim objArg1 As SchGRR
                |           ...
                |          objThisIntf.SetSymbolobjArg1,CatSchIDLCntrSymbolType_Enum

        :param SchGRR i_grr_symbol:
        :param int i_e_symbol_type: enum cat_sch_idl_cntr_symbol_type
        :rtype: None
        """
        return self.sch_grr_cntr.SetSymbol(i_grr_symbol.com_object, i_e_symbol_type)
        # # # # Autogenerated comment: 
        # # some methods require a system service call as the methods expects a vb array object
        # # passed to it and there is no way to do this directly with python. In those cases the following code
        # # should be uncommented and edited accordingly. Otherwise completely remove all this.
        # # vba_function_name = 'set_symbol'
        # # vba_code = """
        # # Public Function set_symbol(sch_grr_cntr)
        # #     Dim iGRRSymbol (2)
        # #     sch_grr_cntr.SetSymbol iGRRSymbol
        # #     set_symbol = iGRRSymbol
        # # End Function
        # # """

        # # system_service = SystemService(self.application.SystemService)
        # # return system_service.evaluate(vba_code, 0, vba_function_name, [self.com_object])

    def __repr__(self):
        return f'SchGRRCntr(name="{self.name}")'
