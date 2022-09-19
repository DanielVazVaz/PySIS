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
        self.update_flowsheet()
        
    def update_flowsheet(self) -> None:
        """
        In case you have created/deleted streams or units, this recalculates the operations, mass streams, and energy streams. 
        """
        dictionary_units = {"Heat Transfer Equipment": HeatExchanger,
                            "Prebuilt Column": DistillationColumn}
        self.Operations     = {str(i):dictionary_units[i.ClassificationName](i) if i.ClassificationName 
                               in dictionary_units else ProcessUnit(i) for i in self.case.Flowsheet.Operations}
        self.MatStreams     = {str(i):MaterialStream(i, self.comp_list) for i in self.case.Flowsheet.MaterialStreams}
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
        Closes the instance and the HYSYS connection. If you do not close it,
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

class ProcessStream:
    def __init__(self, COMObject):
        self.COMObject   = COMObject 
        self.connections = self.get_connections()
        self.name        = self.COMObject.name
        
    def get_connections(self) -> dict:
        upstream   = [i.name for i in self.COMObject.UpstreamOpers]
        downstream = [i.name for i in self.COMObject.DownstreamOpers]
        return {"Upstream": upstream, "Downstream": downstream}
    
class MaterialStream(ProcessStream):
    def __init__(self, COMObject, comp_list: list):
        """"
        Reads the COMObjectfrom the simulation. This class designs a material stream, which has a series
        of properties.
        """
        super().__init__(COMObject)
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
    
    def set_property(self, property_dict: dict) -> None:
        """
        Specify a dictionary with different properties and the units you desire for them.
        Set the properties returned in another dict. 
        
        Each element of the dict must be a tuple with (value:float, units:string). For the
        case of component properties, such as compmassflow, the structure is: (value:dict, unit:string)
        where the value dictionary has the form {"ComponentName": value}.
        
        If some of the properties fail to be set (Normally because the property is not free),
        then it will stop and return an AssertionError. Any property after hte one that fails
        is not set.
        """
        properties_not_found = []
        for property in property_dict:
            if len(property_dict[property]) < 2:
                units = "Adimensional"
            else:
                units = property_dict[property][1]
            value = property_dict[property][0]
            og_property = property
            property = property.upper()
            if property == "PRESSURE":
                self.set_pressure(value, units)
            elif property == "TEMPERATURE":
                self.set_temperature(value, units)
            elif property == "MASSFLOW" or property == "MASS_FLOW":
                self.set_massflow(value, units)
            elif property == "MOLARFLOW" or property == "MOLAR_FLOW":
                self.set_molarflow(value, units)
            elif property == "COMPMASSFLOW" or property == "COMPONENT_MASS_FLOW":
                self.set_compmassflow(value, units)
            elif property == "COMPMOLARFLOW" or property == "COMPONENT_MOLAR_FLOW":
                self.set_compmolarflow(value, units)
            elif property == "COMPMASSFRACTION" or property == "COMPONENT_MASS_FRACTION":
                self.set_compmassfraction(value)
            elif property == "COMPMOLARFRACTION" or property == "COMPONENT_MOLAR_FRACTION":
                self.set_compmolarfraction(value)
            else:
                properties_not_found.append(og_property)
        if properties_not_found:
            print(f"WARNING. The following properties were not found: {properties_not_found}")
            
    def get_pressure(self, units: str = "bar") -> float:
        return self.COMObject.Pressure.GetValue(units)
    
    def set_pressure(self, value: float, units: str = "bar") -> None:
        assert self.COMObject.Pressure.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Pressure.SetValue(-32767.0, "bar")
        else:
            self.COMObject.Pressure.SetValue(value, units)
    
    def get_temperature(self, units:str = "K") -> float:
        return self.COMObject.Temperature.GetValue(units)
    
    def set_temperature(self, value: float, units: str = "K") -> None:
        assert self.COMObject.Temperature.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Temperature.SetValue(-32767.0, "K")
        else: 
            self.COMObject.Temperature.SetValue(value, units)
        
    def get_massflow(self, units: str = "kg/h") -> float:
        return self.COMObject.MassFlow.GetValue(units)
    
    def set_massflow(self, value: float, units:str = "kg/h") -> None:
        assert self.COMObject.MassFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MassFlow.SetValue(-32767.0, "kg/h")
        else:
            self.COMObject.MassFlow.SetValue(value, units)
    
    def get_molarflow(self, units: str = "kgmole/h") -> float:
        return self.COMObject.MolarFlow.GetValue(units)
    
    def set_molarflow(self, value: float, units:str = "kgmole/h") -> None:
        assert self.COMObject.MolarFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MolarFlow.SetValue(-32767.0, "kgmole/h")
        else:
            self.COMObject.MolarFlow.SetValue(value, units)
    
    def get_compmassflow(self, units: str = "kg/h") -> dict:
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFlow.GetValues(units))}
    
    def set_compmassflow(self, values: dict, units: str = "kg/h") -> None:
        assert self.COMObject.ComponentMassFlow.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFlow.SetValues(values_to_set, units)
        
    def get_compmolarflow(self, units: str = "kgmole/h") -> dict:
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFlow.GetValues(units))}
    
    def set_compmolarflow(self, values: dict, units: str = "kgmole/h") -> None:
        assert self.COMObject.ComponentMolarFlow.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFlow.SetValues(values_to_set, units)
        
    def get_compmassfraction(self) -> dict:
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFraction.GetValues(""))}
    
    def set_compmassfraction(self, values: dict) -> None:
        assert self.COMObject.ComponentMassFraction.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFraction.SetValues(values_to_set, "")
        
    def get_compmolarfraction(self) -> dict:
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFraction.GetValues(""))}

    def set_compmolarfraction(self, values: dict) -> None:
        assert self.COMObject.ComponentMolarFraction.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFraction.SetValues(values_to_set, "")

class EnergyStream(ProcessStream):
    def __init__(self, COMObject):
        """"
        Reads the COMObject from the simulation
        """
        super().__init__(COMObject)
        
    def get_power(self, units: str = "kW") -> float:
        return self.COMObject.Power.GetValue(units)
    
    def set_power(self, value: float, units: str = "kW") -> None:
        assert self.COMObject.Power.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Power.SetValue(-32767.0, "kW")
        else:
            self.COMObject.Power.SetValue(value, units)
            

class ProcessUnit:
    def __init__(self, COMObject):
        self.COMObject      = COMObject 
        self.classification = self.COMObject.ClassificationName
        self.name           = self.COMObject.name

class HeatExchanger(ProcessUnit):
    def __init__(self, COMObject):
        super().__init__(COMObject)
        assert self.classification == "Heat Transfer Equipment"
        self.connections = self.get_connections()
        
    def get_connections(self) -> dict:
        feed            = self.COMObject.FeedStream.name
        product         = self.COMObject.ProductStream.name
        energy_stream   = self.COMObject.EnergyStream.name 
        return {"Feed": feed, "Product": product, 
                "Energy stream": energy_stream}
    
    def get_pressuredrop(self, units: str = "bar") -> float:
        return self.COMObject.PressureDrop.GetValue(units)
    
    def set_pressuredrop(self, value: float, units: str = "bar") -> None:
        assert self.COMObject.PressureDrop.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.PressureDrop.SetValue(-32767.0, "bar")
        else:
            self.COMObject.PressureDrop.SetValue(value, units)
            
    def get_deltaT(self, units: str = "K") -> float:
        return self.COMObject.DeltaT.GetValue(units)

    def set_deltaT(self, value: float, units: str = "K"):
        assert self.COMObject.DeltaT.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.DeltaT.SetValue(-32767.0, "K")
        else:
            self.COMObject.DeltaT.SetValue(value, units)
            
class DistillationColumn(ProcessUnit):
    def __init__(self, COMObject):
        super().__init__(COMObject)
        self.column_flowsheet = self.COMObject.ColumnFlowsheet
        try:
            self.main_tower = self.column_flowsheet.Operations["Main Tower"]
        except:
            self.main_tower = "Main Tower could not be located"
            
    def get_feedtrays(self) -> str:
        return [i.name for i in self.main_tower.FeedStages]
    
    def set_feedtray(self, stream: MaterialStream, level: int) -> None:
        self.main_tower.SpecifyFeedLocation(stream, level)
        
    def get_numberstages(self) -> int:
        return self.main_tower.NumberOfTrays
    
    def set_numberstages(self, n:int) -> None:
        self.main_tower.NumberOfTrays = n
        
    def run(self, reset = False) -> None:
        if reset:
            self.column_flowsheet.Reset()
        self.column_flowsheet.Run()
    
    def get_convergence(self) -> bool:
        return self.column_flowsheet.Cfsconverged
    
    def get_specifications(self) -> None:
        return {i.name: {"Active": i.isActive,
                         "Goal": i.Goal.GetValue(),
                         "Current value": i.Current.GetValue()} for i in self.column_flowsheet.specifications}
    
    def set_specification(self, spec: str, value: float, units: str = None) -> None:
        self.column_flowsheet.specifications[spec].Goal.SetValue(value, units)
        
            



