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
        
    def get_property(self, property_dict: dict) -> dict:
        """
        Specify a dictionary with different properties and the units you desire for them.
        Get the properties returned in another dict. 
        """
        result_dict = {}
        properties_not_found = []
        for property in property_dict:
            units = property_dict[property]
            og_property = property
            property = property.upper()
            if property == "PRESSURE":
                result_dict[og_property] = self.get_pressure(units)
            elif property == "TEMPERATURE":
                result_dict[og_property] = self.get_temperature(units)
            elif property == "MASSFLOW" or property == "MASS_FLOW":
                result_dict[og_property] = self.get_massflow(units)
            elif property == "MOLARFLOW" or property == "MOLAR_FLOW":
                result_dict[og_property] = self.get_molarflow(units)
            elif property == "COMPMASSFLOW" or property == "COMPONENT_MASS_FLOW":
                result_dict[og_property] = self.get_compmassflow(units)
            elif property == "COMPMOLARFLOW" or property == "COMPONENT_MOLAR_FLOW":
                result_dict[og_property] = self.get_compmolarflow(units)
            elif property == "MASSFRACTION" or property == "MASS_FRACTION":
                result_dict[og_property] = self.get_compmassfraction()
            elif property == "MOLARFRACTION" or property == "MOLAR_FRACTION":
                result_dict[og_property] = self.get_compmolarfraction()
            else:
                properties_not_found.append(og_property)
        if properties_not_found:
            print(f"WARNING. The following properties were not found: {properties_not_found}")
        return result_dict
            
    def get_pressure(self, units = "bar"):
        return self.COMObject.Pressure.GetValue(units)
    
    def set_pressure(self, value, units = "bar"):
        assert self.COMObject.Pressure.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Pressure.SetValue(-32767.0, "bar")
        else:
            self.COMObject.Pressure.SetValue(value, units)
    
    def get_temperature(self, units = "K"):
        return self.COMObject.Temperature.GetValue(units)
    
    def set_temperature(self, value, units = "K"):
        assert self.COMObject.Temperature.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Temperature.SetValue(-32767.0, "K")
        else: 
            self.COMObject.Temperature.SetValue(value, units)
        
    def get_massflow(self, units = "kg/h") -> float:
        return self.COMObject.MassFlow.GetValue(units)
    
    def set_massflow(self, value: float, units:str = "kg/h") -> None:
        assert self.COMObject.MassFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MassFlow.SetValue(-32767.0, "kg/h")
        else:
            self.COMObject.MassFlow.SetValue(value, units)
    
    def get_molarflow(self, units = "kgmole/h") -> float:
        return self.COMObject.MolarFlow.GetValue(units)
    
    def set_molarflow(self, value: float, units:str = "kgmole/h") -> None:
        assert self.COMObject.MolarFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MolarFlow.SetValue(-32767.0, "kgmole/h")
        else:
            self.COMObject.MolarFlow.SetValue(value, units)
    
    def get_compmassflow(self, units = "kg/h"):
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFlow.GetValues(units))}
    
    def set_compmassflow(self, values: dict, units = "kg/h"):
        assert self.COMObject.ComponentMassFlow.State == 1, "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFlow.SetValues(values_to_set, units)
        
    def get_compmolarflow(self, units = "kgmole/h"):
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFlow.GetValues(units))}
    
    def set_compmolarflow(self, values: dict, units = "kgmole/h"):
        assert self.COMObject.ComponentMolarFlow.State == 1, "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFlow.SetValues(values_to_set, units)
        
    def get_compmassfraction(self):
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFraction.GetValues(""))}
    
    def set_compmassfraction(self, values: dict):
        assert self.COMObject.ComponentMassFraction.State == 1, "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFraction.SetValues(values_to_set, "")
        
    def get_compmolarfraction(self):
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFraction.GetValues(""))}

    def set_compmolarfraction(self, values: dict):
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFraction.SetValues(values_to_set, "")

    def get_connections(self):
        upstream   = [i for i in self.COMObject.UpstreamOpers]
        downstream = [i for i in self.COMObject.DownstreamOpers]
        return {"Upstream": upstream, "Downstream": downstream}
    
    
class EnergyStream:
    def __init__(self, COMObject):
        """"
        Reads the COMObject from the simulation
        """
        self.COMObject = COMObject
            