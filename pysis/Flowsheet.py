#!/usr/bin/env python
# -*- coding: utf-8 -*-

import win32com.client as win32
import os 

class Simulation:
    def __init__(self, path: str) -> None:
        """
        Connects to a given simulation.
        
        Inputs:
            - path: String with the raw path to the HYSYS file. If "Active", chooses the open HYSYS flowsheet.
        """
        self.app = win32.Dispatch("HYSYS.Application")
        if path == "Active":
            self.case = self.app.ActiveDocument
        else:
            file = os.path.abspath(path)
            self.case = self.app.SimulationCases.Open(file)
        self.file_name      = self.case.Title.Value
        self.thermo_package = self.case.Flowsheet.FluidPackage.PropertyPackageName
        self.comp_list      = [i.name for i in self.case.Flowsheet.FluidPackage.Components]
        self.Solver         = self.case.Solver
        self.Operations     = {str(i):i for i in self.case.Flowsheet.Operations}
        self.MatStreams     = {str(i):MaterialStream(i, self.comp_list) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams    = {str(i):EnergyStream(i) for i in self.case.Flowsheet.EnergyStreams}
        
    def update_flowsheet(self) -> None:
        """
        In case you have created/deleted streams or units, this recalculates the operations, mass streams, and energy streams. 
        """
        self.Operations     = {str(i):i for i in self.case.Flowsheet.Operations}
        self.MatStreams     = {str(i):MaterialStream(i) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams    = {str(i):EnergyStream(i) for i in self.case.Flowsheet.EnergyStreams}
    
    def set_visible(self, visibility:int = 0) -> None:
        """
        Sets the visibility of the flowsheet.
        
        Input:
            - visibility: Integer. If 1, it shows the flowsheet. If 0, it keeps it invisible.
        """
        self.case.Visible = visibility
        
    def solver_state(self, state:int = 1) -> None:
        """
        Sets if the solver is active or not.
        
        Input:
            - state: Integer. If 1, the solver is active. If 0, the solver is not active
        """
        self.Solver.CanSolve = state
        
    def close(self) -> None:
        """
        Closes the instance and the HYSYS connection. If you do not close it before working in the flowsheet,
        the task will remain and you will have to stop it from the task manage. 
        """
        self.case.Close()
        self.app.quit()
        del self
    
    def save(self) -> None:
        """
        Saves the current state of the flowsheet. 
        """
        self.case.Save()
        
    def __str__(self) -> str:
        """
        Prints the basic information of the flowsheet.
        """
        return f"File: {self.file_name}\nThermodynamical package: {self.thermo_package}\nComponent list: {self.comp_list}"
    
class MaterialStream:
    def __init__(self, COMObject, comp_list):
        """"
        Reads the COMObjectfrom the simulation. This class designs a material stream, which has a series
        of properties.
        """
        self.COMObject = COMObject
        self.comp_list = comp_list
        
    def get_massflow(self, units = "kg/h") -> float:
        """
        Returns the total mass flow of the stream. The valid units are those that appear in HYSYS.
        
        Input:
            -units: String that indicates the units of the total mass flow.
        """
        return self.COMObject.MassFlow.GetValue(units)
    
    def set_massflow(self, value: float, units:str = "kg/h") -> None:
        if value == "empty":
            self.COMObject.MassFlow.SetValue(-32767.0, "kg/h")
        else:
            self.COMObject.MassFlow.SetValue(value, units)
    
    def get_molarflow(self, units = "kgmole/h") -> float:
        """
        Returns the total molar flow of the stream. The valid units are those that appear in HYSYS.
        
        Input:
            -units: String that indicates the units of the total molar flow.
        """
        return self.COMObject.Molar.GetValue(units)
    
    def set_molarflow(self, value: float, units:str = "kgmole/h") -> None:
        if value == "empty":
            self.COMObject.MolarFlow.SetValue(-32767.0, "kgmole/h")
        else:
            self.COMObject.MolarFlow.SetValue(value, units)
    
    def get_compmassflow(self, units = "kg/h"):
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFlow.GetValues(units))}
    
    def set_compmassflow(self, values: dict, units = "kg/h"):
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFlow.SetValues(values_to_set, units)
    
    def set_compmassfraction(self, values: dict):
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFraction.SetValues(values_to_set, "")
        
    def set_compmolarfraction(self, values: dict):
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFlow.SetValues(values_to_set, "")
        
        
class EnergyStream:
    def __init__(self, COMObject):
        """"
        Reads the COMObject from the simulation
        """
        self.COMObject = COMObject
            