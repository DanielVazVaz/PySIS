# Basic how to

## 1. Make sure that the `pysys` folder can be located
By default, running Python from an IDE such as VSCode adds the folder where script you are running is to the path. However, if from that folder the `pysys` folder is not available, you will have to add it to the path manually. When using interactive sessions, such as Jupyter Notebook, IPython, or others, this path may even vary. As an easy and direct fix, I recommend doing the following:

```python
import sys
sys.path.append(r"PathToTheDirectoryWherePYSYSis")
from pysis.flowsheet import Simulation
```

The worst that can happen if you run this a bunch of times in an interactive interpreter is that you fill the `sys.path` of that session with the same folder. Since it resets each time you close the session, that should not be a problem. I also reccomend that, **unless you really know what you want to do, run it in an interactive environment**, so you can feel safer checking that no weird things are changing. 

## 2. The `Simulation` class
This is the main class that you will be using, since it contains all the information from the flowsheet, albeit only some of it can be accessed easily. In order to load data from a simulation, simply create an instance directing to the path of the hysys file.

```python
Flowsheet = Simulation(path = r"PathToYourFile")
```

It can also read from an already open case.
```python
Flowsheet = Simulation(path = "Active")
```

As a default, when opening a case, it will not be visible. If you want to set it visible, which is useful for checking stuff, but not so much for doing things in a loop, you have to set the attribute `set_visible` to 1.

```python
Flowsheet.set_visible(1)
```

The following are the most important atributes and methods of this instance:
```python
Flowsheet.MatStreams  # Dictionary with all the material streams in the process

Flowsheet.EnerStreams # Dictionary with all the energy streams in the process

Flowsheet.Operations  # Dictionary with all the unitary operations, as well as logical operations, in the flowsheet

Flowsheet.solver_state(solver_state) # Activate with 1, deactivate with 0.

Flowsheet.save() # Saves the flowsheet. Careful with this one. 

print(Flowsheet)      # Indicates information about the loaded flowsheet, in case you are not sure which one is
```

## 3. How to change things in the streams
The basic code to read some attributes from a material stream is:
```python
name_stream = "1"   # Substitute with the name you have in HYSYS

properties_to_read = {
    "Temperature": "K",
    "Pressure": "bar",
    "MassFlow": "kg/h",
    "MolarFlow": "kgmole/h"
}

read_properties = Flowsheet.MatStreams[name_stream].get_property(properties_to_read)
print(read_properties)
```

For now, the properties to read include temperature, pressure, mass and molar total flows, mass and molar component flows, mass and molar fraction. 

To set properties, there are diverse methods.

```python
name_stream = "1"
T = 300
P = 200
x = {"Water": 0.5, "Methane": 0.5}
Flowsheet.MatStreams[name_stream].set_pressure(P, "bar")
Flowsheet.MatStreams[name_stream].set_temperature(T, "K")
Flowsheet.MatStreams[name_stream].set_compmassfraction(x)
```

There are similar methods for the properties that can be read. For now.

## 4. Close
Remember to close when you are finished, or at the end of a script.
```python
Flowsheet.close()
```

If you do not close by code, there will be a process still in the task manager, which I recommend to manually stop before continuing. 