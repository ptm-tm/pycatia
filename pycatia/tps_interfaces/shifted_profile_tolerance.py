#! usr/bin/python3.9
"""
    Module initially auto generated using V5Automation files from CATIA V5 R28 on 2020-09-25 14:34:21.593357

    .. warning::
        The notes denoted "CAA V5 Visual Basic Help" are to be used as reference only.
        They are there as a guide as to how the visual basic / catscript functions work
        and thus help debugging in pycatia.
        
"""

from pycatia.system_interfaces.any_object import AnyObject


class ShiftedProfileTolerance(AnyObject):

    """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357)

                | System.IUnknown
                |     System.IDispatch
                |         System.CATBaseUnknown
                |             System.CATBaseDispatch
                |                 System.AnyObject
                |                     ShiftedProfileTolerance
                | 
                | Interface for accessing shifted tolerance zone informations of a
                | TPS.
    
    """

    def __init__(self, com_object):
        super().__init__(com_object)
        self.shifted_profile_tolerance = com_object

    @property
    def shift_value(self) -> float:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357)
                | o Property ShiftValue() As double
                | 
                |     Retrieves or sets shift value of tolerance zone (in millimeters). The shift
                |     value is the distance between the toleranced surface and the median surface of
                |     tolerance zone. The value is always positive because shift side is given by
                |     GetShiftSide method.

        :return: float
        :rtype: float
        """

        return self.shifted_profile_tolerance.ShiftValue

    @shift_value.setter
    def shift_value(self, value: float):
        """
        :param float value:
        """

        self.shifted_profile_tolerance.ShiftValue = value

    def get_shift_direction(self, op_direction: tuple) -> tuple:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357))
                | o Sub GetShiftDirection(CATSafeArrayVariant opDirection)
                | 
                |     Retrieves the shift direction by two points.
                | 
                |     Parameters:
                | 
                |         opDirection
                |             The first 3 values of opDirection correspond to the X,Y and Z
                |             values of the start point respectively and the next 3 values correspond to the
                |             X, Y and Z values of the end point respectively. 
                | 
                |     Example:
                | 
                |           This example gets the start and end points in a VB
                |           Script
                |          Dim oTab(6) As CATSafeArrayVariant
                |          Set shiftTol = annotation.ShiftedProfileTolerance
                |          shiftTol.GetShiftDirection(oTab)
                |          oStream.Write "      Shifted Direction Start Point : "& oTab(0) & " " & oTab(1) & " " & oTab(2) & sLF
                |          oStream.Write "      Shifted Direction End Point : "& oTab(3) & " " & oTab(4) & " " & oTab(5) & sLF

        :param tuple op_direction:
        :return: tuple
        :rtype: tuple
        """
        return self.shifted_profile_tolerance.GetShiftDirection(op_direction)
        # # # # Autogenerated comment: 
        # # some methods require a system service call as the methods expects a vb array object
        # # passed to it and there is no way to do this directly with python. In those cases the following code
        # # should be uncommented and edited accordingly. Otherwise completely remove all this.
        # # vba_function_name = 'get_shift_direction'
        # # vba_code = """
        # # Public Function get_shift_direction(shifted_profile_tolerance)
        # #     Dim opDirection (2)
        # #     shifted_profile_tolerance.GetShiftDirection opDirection
        # #     get_shift_direction = opDirection
        # # End Function
        # # """

        # # system_service = SystemService(self.application.SystemService)
        # # return system_service.evaluate(vba_code, 0, vba_function_name, [self.com_object])

    def get_shift_side(self, op_point: tuple) -> tuple:
        """
        .. note::
            :class: toggle

            CAA V5 Visual Basic Help (2020-09-25 14:34:21.593357))
                | o Sub GetShiftSide(CATSafeArrayVariant opPoint)
                | 
                |     Retrieves shift side.
                | 
                |     Parameters:
                | 
                |         opPoint
                |             a mathematical point located on the shift side of surface. The 3
                |             values of opPoint correspond to the X,Y and Z values of the point located on
                |             the shift side of surface. 
                | 
                |     Example:
                | 
                |           This example gets the shift side point in a VB
                |           Script
                |          Dim oShiftTab(3) As CATSafeArrayVariant
                |          Set shiftTol = annotation.ShiftedProfileTolerance
                |          shiftTol.GetShiftSide(oShiftTab)
                |          oStream.Write "      Shifted Side Point : "& oShiftTab(0) & " " & oShiftTab(1) & " " & oShiftTab(2) & sLF

        :param tuple op_point:
        :return: tuple
        :rtype: tuple
        """
        return self.shifted_profile_tolerance.GetShiftSide(op_point)
        # # # # Autogenerated comment: 
        # # some methods require a system service call as the methods expects a vb array object
        # # passed to it and there is no way to do this directly with python. In those cases the following code
        # # should be uncommented and edited accordingly. Otherwise completely remove all this.
        # # vba_function_name = 'get_shift_side'
        # # vba_code = """
        # # Public Function get_shift_side(shifted_profile_tolerance)
        # #     Dim opPoint (2)
        # #     shifted_profile_tolerance.GetShiftSide opPoint
        # #     get_shift_side = opPoint
        # # End Function
        # # """

        # # system_service = SystemService(self.application.SystemService)
        # # return system_service.evaluate(vba_code, 0, vba_function_name, [self.com_object])

    def __repr__(self):
        return f'ShiftedProfileTolerance(name="{ self.name }")'
