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
    
    def set_visible(self, visibility = 0):
        self.case.Visible = visibility
        
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
            