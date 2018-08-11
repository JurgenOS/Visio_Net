# Visio_Net
This repo was created to keep the code, that I wrote trying to get a script for automatic topology drawing  in MS Visio application.

The code uses the win32com library.
To make it work I'm going to use the "cdp/lldp neighbor" tables.


### TODO list

- [x] Write the 'node_creator' module to describe the Node class.
- [x] Add the 'connect' method to Node class to link the nodes between each other.
- [x] Write the 'three_tires_topo' module to draw the hierarchy network topoology.
      The hierarchy depends on the amount of the links.
      The more links the hiegher hierarchy.
- [ ] Write the script to parse 'cdp neighbor' files and return the dictionary like this:
```
template = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                   'Fa0/2': {'R6': 'Fa0/0'},
                   'Fa0/3': {'R8': 'Fa0/4'}},
            'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
            'R6': {'Fa0/0': {'R4': 'Fa0/2'},
                   'Fa0/1': {'R8': 'Fa0/3'}},
            'R7': {'Fa0/0': {'R8': 'Fa0/2'}},
            'R8': {'Fa0/2': {'R7': 'Fa0/0'},
                   'Fa0/3': {'R6': 'Fa0/1'},
                   'Fa0/4': {'R4': 'Fa0/3'}}}
```
- [x] Wirte the script 'template_duplicate_deleter' that parses the previous dictionary and deletes the duplicate links.
- [ ] Write the main script that gets the arguments and call the others scripts.
