#!/usr/bin/env python
# -*- coding: utf-8 -*-

import win32com.client as win32
import os 

class Simulation:
    def __init__(self, path: str):
        """
        Connects to the given simulation.
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
        self.MatStreams     = {str(i):MaterialStream(i) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams    = {str(i):i for i in self.case.Flowsheet.EnergyStreams}
        
    def update_flowsheet(self):
        """
        In case you have made changes, this recalculates the operations, mass streams, and energy streams.
        """
        self.Operations     = {str(i):i for i in self.case.Flowsheet.Operations}
        self.MatStreams     = {str(i):MaterialStream(i) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams    = {str(i):EnergyStream(i) for i in self.case.Flowsheet.EnergyStreams}
    
    def set_visible(self, visibility = 0):
        self.case.Visible = visibility
        
    def solver_state(self, state):
        self.Solver.CanSolve = state
        
    def close(self):
        self.case.Close()
        self.app.quit()
        del self
    
    def save(self):
        self.case.Save()
        
    def __str__(self):
        """
        Prints the basic information of the flowsheet.
        """
        return f"File: {self.file_name}\nThermodynamical package: {self.thermo_package}\nComponent list: {self.comp_list}"
    
class MaterialStream:
    def __init__(self, COMObject):
        """"
        Reads the COMObjectfrom the simulation
        """
        self.COMObject = COMObject
        
    def get_massflow(self, units = "kg/h"):
        return self.COMObject.MassFlow.GetValue(units)
    
class EnergyStream:
    def __init__(self, COMObject):
        """"
        Reads the COMObject from the simulation
        """
        self.COMObject = COMObject
            