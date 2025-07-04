import os 

class Simulation:
    """Connects to a given simulation.

    Args:
        path (str): String with the raw path to the HYSYS file. If "Active", chooses the open HYSYS flowsheet.
        version (float): HYSYS version number. For example, for HYSYS version 11, it will be 11.0.. Defaults to None.

    Attributes:
        app: HYSYS application object.
        case: HYSYS simulation case object.
        file_name (str): Name of the HYSYS file.
        thermo_package (str): Name of the thermodynamic package used in the simulation.
        comp_list (list): List of component names in the simulation.
        components (dict): Dictionary of components in the simulation, where keys are component names and values are COMObjects.
        Solver: HYSYS solver object.
        ReactionSets (dict): Dictionary of reaction sets in the simulation.
        Operations (dict): Dictionary of operations in the simulation.
        MatStreams (dict): Dictionary of material streams in the simulation.
        InputMatStreams (dict): Dictionary of input material streams in the simulation.
        OutputMatStreams (dict): Dictionary of output material streams in the simulation.
        EnerStreams (dict): Dictionary of energy streams in the simulation.

    Methods:
        update_flowsheet: Updates the references to the operations, material streams, and energy streams.
        set_visible: Sets the visibility of the flowsheet.
        solver_state: Sets if the solver is active or not.
        close: Closes the instance and the HYSYS connection.
        save: Saves the current state of the flowsheet.
        change_comp_list: Changes the active component list to that of the fluid_package specified.
        __str__: Returns the basic information of the flowsheet.

    """

    def __init__(self, path: str, version:float = None) -> None:  
        import win32com.client as win32   
        if not version:
            self.app = win32.Dispatch("HYSYS.Application")
        else:
            self.app = win32.Dispatch("HYSYS.Application.V"+"{v}".format(v=version))
        if path == "Active":
            self.case = self.app.ActiveDocument
        else:
            file = os.path.abspath(path)
            self.case = self.app.SimulationCases.Open(file)
        self.file_name      = self.case.Title.Value
        self.fluid_packages = {i.name: i for i in self.case.BasisManager.FluidPackages}
        self.thermo_package = self.case.Flowsheet.FluidPackage.PropertyPackageName
        self.comp_list      = [i.name for i in self.case.Flowsheet.FluidPackage.Components]
        self.components     = {i.name:Component(i) for i in self.case.Flowsheet.FluidPackage.Components}
        self.Solver         = self.case.Solver
        self.ReactionSets = {i.name:i for i in self.case.BasisManager.ReactionPackageManager.ReactionSets}
        self.update_flowsheet()
        
    def update_flowsheet(self) -> None:
        """Updates the references to the operations, mass streams, and energy streams in the flowsheet.

        This method should be called in case you have created or deleted streams or units in the flowsheet.
        It updates the references to the operations, mass streams, and energy streams based on the changes made.

        Returns:
            None
        """
        dictionary_units = {"heaterop": HeatExchanger,
                            "coolerop": HeatExchanger,
                            "distillation": DistillationColumn,
                            "pfreactorop": PFR}
        self.Operations      = {str(i):dictionary_units[i.TypeName](i) if i.TypeName 
                               in dictionary_units else ProcessUnit(i) for i in self.case.Flowsheet.Operations}
        self.MatStreams      = {str(i):MaterialStream(i) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams     = {str(i):EnergyStream(i) for i in self.case.Flowsheet.EnergyStreams}
        self.InputMatStreams = {str(i): self.MatStreams[i] for i in self.MatStreams if "FeederBlock_{0}".format(self.MatStreams[i].name) in self.MatStreams[i].get_connections()["Upstream"]}
        self.OutputMatStreams = {str(i): self.MatStreams[i] for i in self.MatStreams if "ProductBlock_{0}".format(self.MatStreams[i].name) in self.MatStreams[i].get_connections()["Downstream"]}
        for stream in self.MatStreams:
            name = self.MatStreams[stream].name
            connections = self.MatStreams[stream].get_connections()
            if "FeederBlock_{0}".format(name) in connections["Upstream"]:
                self.InputMatStreams[name] = self.MatStreams[stream]
            if "ProductBlock_{0}".format(name) in connections["Downstream"]:
                self.OutputMatStreams[name] = self.MatStreams[stream]
    def set_visible(self, visibility:int = 0) -> None:
        """Sets the visibility of the flowsheet.

        Args:
            visibility (int, optional): If 1, it shows the flowsheet. If 0, it keeps it invisible. Defaults to 0.
        """        
        self.case.Visible = visibility
        
    def solver_state(self, state: int = 1) -> None:
        """Sets if the solver is active or not.

        Args:
            state (int, optional): If 1, the solver is active. If 0, the solver is not active. Defaults to 1.
        """        
        self.Solver.CanSolve = state
        
    def close(self) -> None:
        """Closes the instance and the HYSYS connection. If you do not close it,
        the task will remain and you will have to stop it from the task manager.
        """        
        self.case.Close()
        self.app.quit()
        del self
    
    def save(self) -> None:
        """Saves the current state of the flowsheet.

        This method saves the current state of the flowsheet by calling the `Save` method of the `case` object.
        """        
        self.case.Save()

    def change_comp_list(self, fluid_package: "COMObject") -> None:
        """Changes the active component list to that of the fluid_package specified.

        Args:
            fluid_package (COMObject): The COMObject of the fluid package to change to. These can be seen in
            self.fluid_packages
        """

        self.comp_list = [i.name for i in fluid_package.Components]
        
    def __str__(self) -> str:
        """Returns a string representation of the flowsheet.

        This method returns a string that contains the basic information of the flowsheet,
        including the file name, thermodynamical package, and component list.

        Returns:
            str: A string representation of the flowsheet.
        """        
        return f"File: {self.file_name}\nThermodynamical package: {self.thermo_package}\nComponent list: {self.comp_list}"

class Component:
    """Component class that represents a component in the HYSYS simulation.
    
    Args:
        COMObject (COMObject): COMObject of the HYSYS component.

    Attributes:
        COMObject (COMObject): The COMObject of the HYSYS component.
        name (str): The name of the component.
        formula (str): The chemical formula of the component.
        MW (float): The molecular weight of the component.
    
    Methods:
        get_CAS: Gets the CAS number of the component.
        get_Pc: Gets the critical pressure of the component.
        get_Tc: Gets the critical temperature of the component.
        get_bp: Gets the boiling point of the component.
    """
    def __init__(self, COMObject) -> None: 
        self.COMObject = COMObject
        self.name  = self.COMObject.name
        self.formula = self.COMObject.Formula.strip()
        self.MW = self.COMObject.MolecularWeight.GetValue()
    
    def get_CAS(self) -> str:
        """Gets the CAS number of the component. 

        Returns:
            str: The CAS number of the component.
        """     
        return self.COMObject.CAS_Number2
    
    def get_Pc(self, units = "Pa") -> float:
        """Gets the critical pressure of the component.

        Args:
            units (str, optional): Units of the critical pressure. Defaults to "Pa".

        Returns:
            float: The critical pressure of the component in the specified units.
        """        
        return self.COMObject.CriticalPressure.GetValue(units)
    
    def get_Tc(self, units = "K") -> float:
        """Gets the critical temperature of the component.

        Args:
            units (str, optional): Units of the critical temperature. Defaults to "K".

        Returns:
            float: The critical temperature of the component in the specified units.
        """        
        return self.COMObject.CriticalTemperature.GetValue(units)
    
    def get_bp(self, units = "K") -> float:
        """Gets the boiling point of the component.

        Args:
            units (str, optional): Units of the boiling point. Defaults to "K".

        Returns:
            float: The boiling point of the component in the specified units.
        """
        return self.COMObject.NormalBoilingPoint.GetValue(units)
    
    def get_Antoine_parameters(self) -> tuple:
        """Gets the Antoine parameters of the component. The equation is as follows:
        ln(P[kPa]) = a + b/(T[K] + c) + d*ln(T[K]) + e*T^f

        Returns:
            tuple: A tuple containing the Antoine parameters. It has a lot of them, tho, 
            from a to j, even though the equation only uses until f.
        """
        return self.COMObject.AntoineCoeffs.GetValues("")

class ProcessStream:
    """Superclass of all streams in the process.

    Initializes the process stream from a COMObject. Gets the connections and the name of the stream, as well as the COMObject.

    Args:
        COMObject (COMObject): COMObject of the HYSYS process stream.

    Attributes:
        COMObject (COMObject): The COMObject of the HYSYS process stream.
        connections (dict): A dictionary of the upstream and downstream connections.
        name (str): The name of the process stream.
    """    
    def __init__(self, COMObject) -> None:     
        self.COMObject   = COMObject 
        self.connections = self.get_connections()
        self.name        = self.COMObject.name
        
    def get_connections(self) -> dict:
        """Stores the connections of the process stream into a dictionary.

        Returns:
            dict: Returns a dictionary of the upstream and downstream connections.
                The dictionary has two keys: "Upstream" and "Downstream".
                The value for each key is a list of names of the connected operations.
        """        
        upstream   = [i.name for i in self.COMObject.UpstreamOpers]
        downstream = [i.name for i in self.COMObject.DownstreamOpers]
        return {"Upstream": upstream, "Downstream": downstream}
    
class MaterialStream(ProcessStream):
    """Reads the COMObjectfrom the simulation. This class designs a material stream, which has a series
    of properties.

    Args:
        COMObject (COMObject): HYSYS COMObject.

    Methods:
        get_properties: Gets the properties of the material stream.
        set_properties: Sets the properties of the material stream.
        get_pressure: Gets the pressure of the material stream.
        set_pressure: Sets the pressure of the material stream.
        get_temperature: Gets the temperature of the material stream.
        set_temperature: Sets the temperature of the material stream.
        get_massflow: Gets the mass flow of the material stream.
        set_massflow: Sets the mass flow of the material stream.
        get_molarflow: Gets the molar flow of the material stream.
        set_molarflow: Sets the molar flow of the material stream.
        get_compmassflow: Gets the component mass flow of the material stream.
        set_compmassflow: Sets the component mass flow of the material stream.
        get_compmolarflow: Gets the component molar flow of the material stream.
        set_compmolarflow: Sets the component molar flow of the material stream.
        get_compmassfraction: Gets the component mass fraction of the material stream.
        set_compmassfraction: Sets the component mass fraction of the material stream.
        get_compmolarfraction: Gets the component molar fraction of the material stream.
        set_compmolarfraction: Sets the component molar fraction of the material stream.
    """ 
    def __init__(self, COMObject):       
        super().__init__(COMObject)
        self.components = {i.name:Component(i) for i in self.COMObject.FluidPackage.Components}
        self.comp_list = list(self.components.keys())
        
    def get_properties(self, property_dict: dict) -> dict:
        """Specify a dictionary with different properties and the units you desire for them.
        Get the properties returned in another dict.

        Args:
            property_dict (dict): Dictionary of properties to get. The format is {"PropertyName":"Units"}. For properties
            without units, i.e., component molar flow and component mass flow, an empty string can be used. The valid properties are
            "PRESSURE", "TEMPERATURE", "MASSFLOW", "MOLARFLOW", "COMPMASSFLOW", "COMPMOLARFLOW", "COMPMASSFRACTION", "COMPMOLARFRACTION"

        Returns:
            dict: Dictionary where the keys are the properties, and the elements are the values read. 
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
            elif property == "COMPMASSFRACTION" or property == "COMPONENT_MASS_FRACTION":
                result_dict[og_property] = self.get_compmassfraction()
            elif property == "COMPMOLARFRACTION" or property == "COMPONENT_MOLAR_FRACTION":
                result_dict[og_property] = self.get_compmolarfraction()
            else:
                properties_not_found.append(og_property)
        if properties_not_found:
            valid_properties = ["PRESSURE", "TEMPERATURE", "MASSFLOW", "MOLARFLOW", "COMPMASSFLOW", "COMPMOLARFLOW", "COMPMASSFRACTION", "COMPMOLARFRACTION"]
            print(f"WARNING. The following properties were not found: {properties_not_found}\nThe valid properties are: {valid_properties}")
        return result_dict
    
    def set_properties(self, property_dict: dict) -> None:
        """Specify a dictionary with different properties and the units you desire for them.
        Set the properties returned in another dict. If some of the properties fail to be set 
        (Normally because the property is not free), then it will stop and return an AssertionError. 
        Any property after the one that fails is not set.

        Args:
            property_dict (dict): Each element of the dict must be a tuple with (value:float, units:string). For the
            case of component properties, such as compmassflow, the structure is: (value:dict, unit:string)
            where the value dictionary has the form {"ComponentName": value}. The valid properties are
            "PRESSURE", "TEMPERATURE", "MASSFLOW", "MOLARFLOW", "COMPMASSFLOW", "COMPMOLARFLOW", "COMPMASSFRACTION", "COMPMOLARFRACTION"
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
            valid_properties = ["PRESSURE", "TEMPERATURE", "MASSFLOW", "MOLARFLOW", "COMPMASSFLOW", "COMPMOLARFLOW", "COMPMASSFRACTION", "COMPMOLARFRACTION"]
            print(f"WARNING. The following properties were not found: {properties_not_found}\nThe valid properties are: {valid_properties}")
            
    def get_pressure(self, units: str = "bar") -> float:
        """Read the pressure.

        Args:
            units (str, optional): Units of the read output. Defaults to "bar".

        Returns:
            float: Value of the output in the specified units in the process stream.
        """        
        return self.COMObject.Pressure.GetValue(units)
    
    def set_pressure(self, value: float, units: str = "bar") -> None:
        """Sets the pressure of the material stream.

        Args:
            value (float): Value to which the pressure is set.
            units (str, optional): Units of the pressure. Defaults to "bar".
        """        
        assert self.COMObject.Pressure.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Pressure.SetValue(-32767.0, "bar")
        else:
            self.COMObject.Pressure.SetValue(value, units)
    
    def get_temperature(self, units:str = "K") -> float:
        """Reads the temperature.

        Args:
            units (str, optional): Units of the read output. Defaults to "K".

        Returns:
            float: Value of the output in the specified units in the process stream.
        """        
        return self.COMObject.Temperature.GetValue(units)
    
    def set_temperature(self, value: float, units: str = "K") -> None:
        """Sets the temperature of the material stream.

        Args:
            value (float): Value to which the temperature is set.
            units (str, optional): Units of the temperature. Defaults to "K".
        """        
        assert self.COMObject.Temperature.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Temperature.SetValue(-32767.0, "K")
        else: 
            self.COMObject.Temperature.SetValue(value, units)
        
    def get_massflow(self, units: str = "kg/h") -> float:
        """Reads the mass flow.

        Args:
            units (str, optional): Units of the read output. Defaults to "kg/h".

        Returns:
            float: Value of the output in the specified units.
        """        
        return self.COMObject.MassFlow.GetValue(units)
    
    def set_massflow(self, value: float, units:str = "kg/h") -> None:
        """Sets the mass flow.

        Args:
            value (float): Value of the mass flow.
            units (str, optional): Units of the mass flow. Defaults to "kg/h".
        """        
        assert self.COMObject.MassFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MassFlow.SetValue(-32767.0, "kg/h")
        else:
            self.COMObject.MassFlow.SetValue(value, units)
    
    def get_molarflow(self, units: str = "kgmole/h") -> float:
        """Reads the molar flow rate.

        Args:
            units (str, optional): Units of the read output. Defaults to "kgmole/h".

        Returns:
            float: Value of the output in the specified units.
        """        
        return self.COMObject.MolarFlow.GetValue(units)
    
    def set_molarflow(self, value: float, units:str = "kgmole/h") -> None:
        """Sets the mole flow.

        Args:
            value (float): Value of the molar flow rate.
            units (str, optional): Units of the molar flow rate. Defaults to "kgmole/h".
        """        
        assert self.COMObject.MolarFlow.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.MolarFlow.SetValue(-32767.0, "kgmole/h")
        else:
            self.COMObject.MolarFlow.SetValue(value, units)
    
    def get_compmassflow(self, units: str = "kg/h") -> dict:
        """Reads the component mass flow.

        Args:
            units (str, optional): Units of the read output. Defaults to "kg/h".

        Returns:
            dict: Dictionary with the species as the keys and the values as the read outputs.
        """        
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFlow.GetValues(units))}
    
    def set_compmassflow(self, values: dict, units: str = "kg/h") -> None:
        """Sets the component mass flow. 

        Args:
            values (dict): The format of the dictionary shall be {"Component": Value}
            units (str, optional): Unit of the component mass flow. Defaults to "kg/h".
        """        
        assert self.COMObject.ComponentMassFlow.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFlow.SetValues(values_to_set, units)
        
    def get_compmolarflow(self, units: str = "kgmole/h") -> dict:
        """Reads the component molar flow rate.

        Args:
            units (str, optional): Units of the read output. Defaults to "kgmole/h".

        Returns:
            dict: A dictionary of all the components and their values in the chosen units.
        """        
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFlow.GetValues(units))}
    
    def set_compmolarflow(self, values: dict, units: str = "kgmole/h") -> None:
        """Sets the component molar flow rate.

        Args:
            values (dict): The format of the dictionary shall be {"Component": Value}.
            units (str, optional): Units of the component molar flow rate. Defaults to "kgmole/h".
        """        
        assert self.COMObject.ComponentMolarFlow.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFlow.SetValues(values_to_set, units)
        
    def get_compmassfraction(self) -> dict:
        """Reads the component mass fraction.

        Returns:
            dict: A dictionary of the components and their values. 
        """        
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMassFraction.GetValues(""))}
    
    def set_compmassfraction(self, values: dict) -> None:
        """Sets the component mass fraction.

        Args:
            values (dict): The format of the dictionary shall be {"Component": Value}
        """        
        assert self.COMObject.ComponentMassFraction.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMassFraction.SetValues(values_to_set, "")
        
    def get_compmolarfraction(self) -> dict:
        """Reads the component molar fraction.

        Returns:
            dict: A dictionary of the components and their values.
        """        
        return {i:j for (i,j) in zip(self.comp_list, self.COMObject.ComponentMolarFraction.GetValues(""))}

    def set_compmolarfraction(self, values: dict) -> None:
        """Sets the component molar fraction.

        Args:
            values (dict): The format of the dictionary shall be {"Component": Value}
        """        
        assert self.COMObject.ComponentMolarFraction.State == (1,)*len(self.comp_list), "The variable is calculated. Cannot modify it."
        values_to_set = [0]*len(self.comp_list)
        for component in values:
            values_to_set[self.comp_list.index(component)] = values[component]
        values_to_set = tuple(values_to_set)
        self.COMObject.ComponentMolarFraction.SetValues(values_to_set, "")

class EnergyStream(ProcessStream):
    """Reads the COMObject from the simulation. This class designs an energy stream.

    Args:
        COMObject (COMObject): Object from the HYSYS simulation.
    
    Methods:
        get_power: Gets the power of the energy stream.
        set_power: Sets the power of the energy stream.
    """      
    def __init__(self, COMObject):  
        super().__init__(COMObject)
        
    def get_power(self, units: str = "kW") -> float:
        """Reads the power/duty of the stream.

        Args:
            units (str, optional): Units of the power. Defaults to "kW".

        Returns:
            float: Value of the power in the specified units.
        """        
        return self.COMObject.Power.GetValue(units)
    
    def set_power(self, value: float, units: str = "kW") -> None:
        """Sets the power of the energy stream.

        Args:
            value (float): Value to set.
            units (str, optional): Units of the power. Defaults to "kW".
        """        
        assert self.COMObject.Power.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.Power.SetValue(-32767.0, "kW")
        else:
            self.COMObject.Power.SetValue(value, units)
            

class ProcessUnit:
    """Superclass for all process units in the flowsheet. Contains the classification of the unit,
    the COMObject, and the name.

    Args:
        COMObject (COMObject): COMObject from HYSYS.

    Attributes:
        COMObject (COMObject): COMObject from HYSYS.
        classification (str): Classification of the process unit.
        name (str): Name of the process unit.
    """        
    def __init__(self, COMObject):
        self.COMObject      = COMObject 
        self.classification = self.COMObject.ClassificationName
        self.name           = self.COMObject.name

class HeatExchanger(ProcessUnit):
    """Child of process unit. Includes the connections of the heat exchanger.

    Args:
        COMObject (COMObject): COMObject from HYSYS.

    Attributes:
        connections (dict): Dictionary with the feed, product, and energy stream.

    Methods:
        get_connections: Reads the connections of the process unit.
        get_pressuredrop: Reads the pressure drop of the unit.
        set_pressuredrop: Sets the pressure drop of the unit.
        get_deltaT: Reads the deltaT of the unit.
        set_deltaT: Sets the deltaT of the unit.
    """
    def __init__(self, COMObject):        
        super().__init__(COMObject)
        assert self.classification == "Heat Transfer Equipment"
        self.connections = self.get_connections()
        
    def get_connections(self) -> dict:
        """Reads the connections of the process unit.

        Returns:
            dict: Dictionary that contains the feed, product and energy stream attached to the process unit.
        """        
        feed            = self.COMObject.FeedStream.name
        product         = self.COMObject.ProductStream.name
        energy_stream   = self.COMObject.EnergyStream.name 
        return {"Feed": feed, "Product": product, 
                "Energy stream": energy_stream}
    
    def get_pressuredrop(self, units: str = "bar") -> float:
        """Reads the pressure drop of the unit.

        Args:
            units (str, optional): Units of the output. Defaults to "bar".

        Returns:
            float: Value of the pressure drop in the specified units.
        """        
        return self.COMObject.PressureDrop.GetValue(units)
    
    def set_pressuredrop(self, value: float, units: str = "bar") -> None:
        """Sets the pressure drop of the unit.

        Args:
            value (float): Value to set. If "empty", it frees the degree of freedom.
            units (str, optional): Units of the pressure drop. Defaults to "bar".
        """        
        assert self.COMObject.PressureDrop.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.PressureDrop.SetValue(-32767.0, "bar")
        else:
            self.COMObject.PressureDrop.SetValue(value, units)
            
    def get_deltaT(self, units: str = "K") -> float:
        """Read the deltaT of the unit.

        Args:
            units (str, optional): Units of the output. Defaults to "K".

        Returns:
            float: Value of the deltaT.
        """        
        return self.COMObject.DeltaT.GetValue(units)

    def set_deltaT(self, value: float, units: str = "K"):
        """Sets the deltaT of the unit. Normally, it is recommended to directly set the temperature
        in the input and output stream. But you do you.

        Args:
            value (float): Value of the deltaT. If "empty", it frees the degree of freedom.
            units (str, optional): Units of the deltaT. Defaults to "K".
        """        
        assert self.COMObject.DeltaT.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            self.COMObject.DeltaT.SetValue(-32767.0, "K")
        else:
            self.COMObject.DeltaT.SetValue(value, units)
            
class DistillationColumn(ProcessUnit):
    """Child of process unit. It contains the COMObject, the ColumnFlowsheet, and the
    Main Tower. In some thermodynamic packages, it seems that the main tower is not available.

    Args:
        COMObject (COMObject): COMObject from HYSYS.
    
    Attributes:
        column_flowsheet (COMObject): COMObject of the column flowsheet.
        main_tower (COMObject): COMObject of the main tower.

    Methods:
        get_feedtrays: Gets the feedtrays of the column.
        set_feedtray: Sets the feedtray of the column.
        get_numberstages: Gets the number of stages of the column.
        set_numberstages: Sets the number of stages of the column.
        run: Runs the column subflowsheet.
        get_convergence: Checks the convergence of the column.
        get_specifications: Reads the specifications of the column.
        set_specifications: Sets the specifications of the column.
    """ 
    def __init__(self, COMObject):       
        super().__init__(COMObject)
        self.column_flowsheet = self.COMObject.ColumnFlowsheet
        try:
            self.main_tower = self.column_flowsheet.Operations["Main Tower"]
        except:
            self.main_tower = "Main Tower could not be located"
            
    def get_feedtrays(self) -> str:
        """Gets the feedtray.

        Returns:
            str: Feedtray.
        """        
        return [i.name for i in self.main_tower.FeedStages]
    
    def set_feedtray(self, stream: str, level: int) -> None:
        """Sets the feedtray. If the solver is deactivated, you will not see the change. Also,
        it unconverges the column. Remember to run() it afterwards.

        Args:
            stream (string): Name of the stream that is already in the column.
            level (int): Level of the column where the feed is set.
        """
        position_feeds = [n for n,i in enumerate(self.main_tower.AttachedFeeds) if i.name == stream][0]
        
        self.main_tower.SpecifyFeedLocation(self.main_tower.AttachedFeeds[position_feeds], level)
        
    def get_numberstages(self) -> int:
        """Reads the number of stages.

        Returns:
            int: Number of stages.
        """        
        return self.main_tower.NumberOfTrays
    
    def set_numberstages(self, n:int) -> None:
        """Sets the number of stages.

        Args:
            n (int): Number of stages.
        """        
        self.main_tower.NumberOfTrays = n
        
    def run(self, reset = False) -> None:
        """Runs the column subflowsheet.

        Args:
            reset (bool, optional): If True, resets the column before running it. Defaults to False.
        """        
        if reset:
            self.column_flowsheet.Reset()
        self.column_flowsheet.Run()
    
    def get_convergence(self) -> bool:
        """Checks the convergence of the column. Careful, if you check after changing the 
        input, such as the feedtray, it does not change from the previous state. Check always after
        a run()

        Returns:
            bool: True if the column is converged. False otherwise.
        """        
        return self.column_flowsheet.Cfsconverged
    
    def get_specifications(self) -> dict:
        """Reads the specifications of the column.

        Returns:
            dict: Dictionary with the specifications, if they are active, their goal value and their current value.
        """        
        return {i.name: {"Active": i.isActive,
                         "Goal": i.Goal.GetValue(),
                         "Current value": i.Current.GetValue()} for i in self.column_flowsheet.specifications}
    
    def set_specifications(self, spec: str, value: float, units: str = None) -> None:
        """Sets the goal value of a specification.

        Args:
            spec (str): Name of the specification.
            value (float): Value to which the goal value is set.
            units (str, optional): Units of the specification. Defaults to None.
        """        
        self.column_flowsheet.specifications[spec].Goal.SetValue(value, units)

class PFR(ProcessUnit):
    """Plug flow reactor. 

    Args:
        COMObject (COMObject): COMObject from HYSYS

    Attributes:
        connections (dict): Dictionary with the feed, product, and energy stream.
        reactions (dict): Dictionary with the reactions and their values.

    Methods:
        get_connections: Gets the connections of the reactor.
        modify_feed: Modifies the feed streams. Adds or removes a bunch of feed streams.
        modify_product: States which is the product stream. It cannot be left empty as far as I know.
        get_reactionset: Reads the name of the reaction set.
        set_reactionset: Sets the reaction set to use.
        get_volume: Returns the volume of the unit.
        set_volume: Sets the volume of the unit.
        get_length: Returns the length of the unit.
        set_length: Sets the length of the unit.
        get_diameter: Returns the diameter of the unit.
        set_diameter: Sets the diameter of the unit.
        get_ntubes: Return the number of tubes.
        set_ntubes: Sets the number of tubes.
        get_voidfraction: Return the void fraction.
        set_voidfraction: Sets the void fraction.
        get_voidvolume: Return the void volume.
        set_voidvolume: Sets the void volume.
        get_properties: Specify a dictionary with different properties and the units you desire for them. Get the properties returned in another dict.
        set_properties: Specify a dictionary with different properties and the units you desire for them. Set the properties returned in another dict.
    """    
    def __init__(self, COMObject):
        super().__init__(COMObject)
        self.connections = self.get_connections()
        
    def get_connections(self) -> dict:
        """Get the connections of the reactor

        Returns:
            dict: Returns the connections
        """        
        feed = [i.name for i in self.COMObject.Feeds]
        product = self.COMObject.Product.name
        try:
            energy = self.COMObject.EnergyStream.name
        except:
            energy = None
        return {"Feed": feed, "Product": product, "EnergyStream": energy}
    
    def modify_feed(self, stream_group: list, movement: str) -> None:
        """Modifies the feed streams. Adds or removes a bunch of feed streams

        Args:
            stream_group (list): List of names of the feed streams to modify 
            movement (str): Either "Add" or "Remove"
        """
        movement = movement.upper()
        for stream in stream_group:
            if movement == "ADD":
                self.COMObject.Feeds.Add(stream)
            elif movement == "REMOVE":
                self.COMObject.Feeds.Remove(stream)
            else:
                print("Movement not valid. Choose either ADD or REMOVE")
                
    def modify_product(self, product_stream: "COMObject") -> None:
        """States which is the product stream. It cannot be left empty
        as far as I know. 

        Args:
            product_stream (COMObject): COMObject which will be the product. You can get
            it from the flowsheet object, MatStreams dictionary, name of the stream, COMObject.
            For example: FS.MatStreams["Stream1"].COMObject
        """
        self.COMObject.Product = product_stream
        
    def get_reactionset(self) -> str:
        """Reads the name of the reaction set

        Returns:
            str: Name of the reaction set.
        """
        return self.COMObject.ReactionSet.name
    
    def set_reactionset(self, rxn_set: "COMObject") -> None:
        """Sets the reaction set to use

        Args:
            rxn_set (COMObject): COMObject of the reaction set. Tends to be in
            the flowsheet object, FS.ReactionSets dictionary. Use the name to access it.
            Consider that the reaction set must be of a compatible type
        """       
        self.COMObject.ReactionSet = rxn_set 

    def get_volume(self, units:str = "m3") -> float:
        """Returns the volume of the unit

        Args:
            units (str, optional): Defaults to "m3".

        Returns:
            float: Volume in the specified units
        """        
        return self.COMObject.TotalVolume.GetValue(units)
    def set_volume(self, value:float, units:str = "m3") -> None:
        """Sets the volume of the unit

        Args:
            value (float): Value to set
            units (str, optional): Defaults to "m3".
        """        
        assert self.COMObject.TotalVolume.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TotalVolume.SetValue(value, units)
    def get_length(self, units: str = "m") -> float:
        """Returns the length of the unit

        Args:
            units (str, optional): Defaults to "m".

        Returns:
            float: Length of the tubes in the specified units
        """        
        return self.COMObject.TubeLength.GetValue(units)
    def set_length(self, value:float, units:str = "m") -> None:
        """Set the length of the unit

        Args:
            value (float): Value to set
            units (str, optional): Defaults to "m".
        """        
        assert self.COMObject.TubeLength.State ==1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TubeLength.SetValue(value, units)
    def get_diameter(self, units:str = "m") -> float:
        """Returns the diameter of the unit

        Args:
            units (str, optional): Defaults to "m".

        Returns:
            float: Diameter of the tubes in the specified units
        """        
        return self.COMObject.TubeDiameter.GetValue(units)
    def set_diameter(self, value:float, units:str = "m"):
        """Set the diameter of the unit

        Args:
            value (float): Value to set
            units (str, optional): Defaults to "m".
        """        
        assert self.COMObject.TubeDiameter.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TubeDiameter.SetValue(value, units)
    def get_ntubes(self) -> int:
        """Return the number of tubes

        Returns:
            float: Return the number of tubes
        """        
        return self.COMObject.NumberOfTubes
    def set_ntubes(self, value:int) -> None:
        """Set the number of tubes

        Args:
            value (int): Value to set
        """        
        if value == "empty":
            value = -32767.0
        self.COMObject.NumberOfTubes = value
    def get_voidfraction(self) -> float:
        """Return the void fraction

        Returns:
            float: Void fraction
        """        
        return self.COMObject.VoidFraction
    def set_voidfraction(self, value:float) -> None:
        """Set the void fraction

        Args:
            value (float): Value to set
        """        
        if value == "empty":
            value = -32767.0
        self.COMObject.VoidFraction = value
    def get_voidvolume(self, units = "m3") -> float:
        """Return the void volume

        Args:
            units (str, optional): Defaults to "m3".

        Returns:
            float: Void volume in the specified units
        """        
        return self.COMObject.VoidVolume.GetValue(units)
    def set_voidvolume(self, value:float, units:str = "m3") -> None:
        """Sets the void volume

        Args:
            value (float): Value to set
            units (str, optional): Defaults to "m3".
        """        
        assert self.COMObject.VoidVolume.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.VoidVolume.SetValue(value, units)
        
    def get_properties(self, property_dict:dict) -> dict:
        """Get a number of properties. Allowed keys are VOLUME, REACTIONSET, TUBELENGTH, TUBEDIAMETER,
        NUMBEROFTUBES, VOIDFRACTION, and VOIDVOLUME.

        Args:
            prop_dict (dict): Dictionary of the desired property. Each element has to have a string indicating the units. For the
            properties that do not have a unit, i.e., VOIDFRACTION and NUMBEROFTUBES, leave the unit an empty string, or None.

        Returns:
            dict: Dictionary with the properties as keys and the value in the specified units as value
        """
        result_dict = {}
        properties_not_found = []
        unitless_properties = ["REACTIONSET", "VOIDFRACTION", "NUMBEROFTUBES"]
        function_property_dict = {"VOLUME": self.get_volume, "REACTIONSET": self.get_reactionset, "TUBELENGTH": self.get_length,
                                  "TUBEDIAMETER": self.get_diameter, "NUMBEROFTUBES": self.get_ntubes, "VOIDFRACTION": self.get_voidfraction, 
                                  "VOIDVOLUME": self.get_voidvolume}
        for property in property_dict:
            units =  property_dict[property]
            og_property = property 
            property = property.upper()
            if property in function_property_dict:
                if property not in unitless_properties:
                    result_dict[og_property] = function_property_dict[property](units)
                else:
                    result_dict[og_property] = function_property_dict[property]()
            else:
                properties_not_found.append(og_property)
                
        if properties_not_found:
                print(f"WARNING. The following properties were not found: {properties_not_found}\nThe valid properties are: {list(function_property_dict.keys())}")
        
        return result_dict
    
    def set_properties(self, property_dict:dict) -> None:
        """Set a number of properties. Allowed keys are VOLUME, REACTIONSET, TUBELENGTH, TUBEDIAMETER,
        NUMBEROFTUBES, VOIDFRACTION, and VOIDVOLUME.

        Args:
            prop_dict (dict): Dictionary of the desired property. Each element has to have a tuple indicating the value and units, i.e., (value,units). 
            For the properties that do not have a unit, i.e., VOIDFRACTION and NUMBEROFTUBES, use an empty string or None as the unit.
        """   
        result_dict = {}
        properties_not_found = []
        unitless_properties = ["REACTIONSET", "VOIDFRACTION", "NUMBEROFTUBES"]
        function_property_dict = {"VOLUME": self.set_volume, "REACTIONSET": self.set_reactionset, "TUBELENGTH": self.set_length,
                                  "TUBEDIAMETER": self.set_diameter, "NUMBEROFTUBES": self.set_ntubes, "VOIDFRACTION": self.set_voidfraction, 
                                  "VOIDVOLUME": self.set_voidvolume}
        for property in property_dict:
            units =  property_dict[property][1]
            value = property_dict[property][0]
            og_property = property 
            property = property.upper()
            if property in function_property_dict:
                if property not in unitless_properties:
                    result_dict[og_property] = function_property_dict[property](value, units)
                else:
                    result_dict[og_property] = function_property_dict[property](value)
            else:
                properties_not_found.append(og_property)
                
        if properties_not_found:
                print(f"WARNING. The following properties were not found: {properties_not_found}\nThe valid properties are: {list(function_property_dict.keys())}")

        
        
    
        


