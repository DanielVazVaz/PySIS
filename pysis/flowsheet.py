import os 

class Simulation:
    """Connects to a given simulation.

    Args:
        path (str): String with the raw path to the HYSYS file. If "Active", chooses the open HYSYS flowsheet.
    """
    def __init__(self, path: str) -> None:  
        import win32com.client as win32     
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
        self.ReactionSets = {i.name:i for i in self.case.BasisManager.ReactionPackageManager.ReactionSets}
        self.update_flowsheet()
        
    def update_flowsheet(self) -> None:
        """In case you have created/deleted streams or units, this recalculates the operations, mass streams, and energy streams.
        """        
        dictionary_units = {"heaterop": HeatExchanger,
                            "coolerop": HeatExchanger,
                            "distillation": DistillationColumn,
                            "pfreactorop": PFR}
        self.Operations     = {str(i):dictionary_units[i.TypeName](i) if i.TypeName 
                               in dictionary_units else ProcessUnit(i) for i in self.case.Flowsheet.Operations}
        self.MatStreams     = {str(i):MaterialStream(i, self.comp_list) for i in self.case.Flowsheet.MaterialStreams}
        self.EnerStreams    = {str(i):EnergyStream(i) for i in self.case.Flowsheet.EnergyStreams}
    
    def set_visible(self, visibility:int = 0) -> None:
        """Sets the visibility of the flowsheet.

        Args:
            visibility (int, optional): If 1, it shows the flowsheet. If 0, it keeps it invisible.. Defaults to 0.
        """        
        self.case.Visible = visibility
        
    def solver_state(self, state:int = 1) -> None:
        """Sets if the solver is active or not.

        Args:
            state (int, optional): If 1, the solver is active. If 0, the solver is not active. Defaults to 1.
        """        
        self.Solver.CanSolve = state
        
    def close(self) -> None:
        """Closes the instance and the HYSYS connection. If you do not close it,
        the task will remain and you will have to stop it from the task manage.
        """        
        self.case.Close()
        self.app.quit()
        del self
    
    def save(self) -> None:
        """Saves the current state of the flowsheet.
        """        
        self.case.Save()
        
    def __str__(self) -> str:
        """Prints the basic information of the flowsheet.

        Returns:
            str: Information of the flowsheet. 
        """        
        return f"File: {self.file_name}\nThermodynamical package: {self.thermo_package}\nComponent list: {self.comp_list}"

class ProcessStream:
    """Superclass of all streams in the process.Initializes the process stream from a COMObject. Gets the connections and the name of the stream, as
    well as the COMObject.

    Args:
        COMObject (COMObject): COMObject of the HYSYS process stream.
    """    
    def __init__(self, COMObject) -> None:     
        self.COMObject   = COMObject 
        self.connections = self.get_connections()
        self.name        = self.COMObject.name
        
    def get_connections(self) -> dict:
        """Stores the connections of the process stream into a dictionary.

        Returns:
            dict: Returns a dictionary of the upstream and downstream connections
        """        
        upstream   = [i.name for i in self.COMObject.UpstreamOpers]
        downstream = [i.name for i in self.COMObject.DownstreamOpers]
        return {"Upstream": upstream, "Downstream": downstream}
    
class MaterialStream(ProcessStream):
    """Reads the COMObjectfrom the simulation. This class designs a material stream, which has a series
    of properties.

    Args:
        COMObject (COMObject): HYSYS COMObject.
        comp_list (list): List of components in the process stream.
    """ 
    def __init__(self, COMObject, comp_list: list):       
        super().__init__(COMObject)
        self.comp_list = comp_list
        
    def get_properties(self, property_dict: dict) -> dict:
        """Specify a dictionary with different properties and the units you desire for them.
        Get the properties returned in another dict.

        Args:
            property_dict (dict): Dictionary of properties to get. The format is {"PropertyName":"Units"}. For properties
            without units, i.e., component molar flow and component mass flow, an empty string can be used.

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
            where the value dictionary has the form {"ComponentName": value}.
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
    """        
    def __init__(self, COMObject):
        self.COMObject      = COMObject 
        self.classification = self.COMObject.ClassificationName
        self.name           = self.COMObject.name

class HeatExchanger(ProcessUnit):
    """Child of process unit. Includes the connections of the heat exchanger.

    Args:
        COMObject (COMObject): COMObject from HYSYS.
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
    
    def set_feedtray(self, stream: MaterialStream, level: int) -> None:
        """Sets the feedtray.

        Args:
            stream (MaterialStream): Stream that goes into the feed tray.
            level (int): Level of the column where the feed is set.
        """        
        self.main_tower.SpecifyFeedLocation(stream, level)
        
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
        """Checks the convergence of the column.

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
        energy  = self.COMObject.EnergyStream.name
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
            the flowsheet object, FS.ReactionSets dictionary. Use the name to access it
        """       
        self.COMObject.ReactionSet = rxn_set 

    def get_volume(self, units:str = "m3") -> float:
        return self.COMObject.TotalVolume.GetValue(units)
    def set_volume(self, value:float, units:str = "m3") -> None:
        assert self.COMObject.TotalVolume.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TotalVolume.SetValue(value, units)
    def get_length(self, units: str = "m") -> float:
        return self.COMObject.TubeLength.GetValue(units)
    def set_length(self, value:float, units:str = "m") -> None:
        assert self.COMObject.TubeLength.State ==1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TubeLength.SetValue(value, units)
    def get_diameter(self, units:str = "m") -> float:
        return self.COMObject.TubeDiameter.GetValue(units)
    def set_diameter(self, value:float, units:str = "m"):
        assert self.COMObject.TubeDiameter.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.TubeDiameter.SetValue(value, units)
    def get_ntubes(self) -> float:
        return self.COMObject.NumberOfTubes
    def set_ntubes(self, value:int) -> None:
        if value == "empty":
            value = -32767.0
        self.COMObject.NumberOfTubes = value
    def get_voidfraction(self) -> float:
        return self.COMObject.VoidFraction
    def set_voidfraction(self, value:float) -> None:
        if value == "empty":
            value = -32767.0
        self.COMObject.VoidFraction = value
    def get_voidvolume(self, units = "m3") -> float:
        return self.COMObject.VoidVolume.GetValue(units)
    def set_voidvolume(self, value:float, units:str = "m3") -> None:
        assert self.COMObject.VoidVolume.State == 1, "The variable is calculated. Cannot modify it."
        if value == "empty":
            value = -32767.0
        self.COMObject.VoidVolume.SetValue(value, units)
    
        


