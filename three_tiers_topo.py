# -*- encoding: utf-8 -*-

def three_tiers_topo(cdp_neigh_tempalte):
    '''
    build 3 tiers in the page according to amount of links
    the more links the higher position
    !!! there are only 3 tiers !!! 
    '''
    templ = deleter.templ_parser(cdp_neigh_tempalte)
    print('templ ', templ)
    tire_1 = {}
    tire_2 = {}
    tire_3 = {}
    
    for node, value in cdp_neigh_tempalte.items():
    
        if len(value) >= 3:

            tire_1[node] = node_creator.Node(node, 24)
            tire_1[node].create(MasterFile, pagObj, 'L2_24_Blue', appVisio, x=len(tire_1)*2, y=5)

        if 3 > len(value) >= 2:

            tire_2[node] = node_creator.Node(node, 24)
            tire_2[node].create(MasterFile, pagObj, 'L2_24_Blue', appVisio, x=len(tire_2)*2, y=2)

        if 2 > len(value) >= 1:

            tire_3[node] = node_creator.Node(node, 24)
            tire_3[node].create(MasterFile, pagObj, 'L2_24_Blue', appVisio, x=len(tire_3)*2, y=0)
        #print('node: ', node, '\n',
        #      'value: ', value)



    nodes = {}
    nodes.update(tire_1)
    nodes.update(tire_2)    
    nodes.update(tire_3)

    #print('nodes ', nodes)

    		
    for key, value in templ.items():
        for interf, item in value.items():

            neigh_name, neigh_intf = list(item.items())[0]
            print('\n', '#########################', '\n')
            print('Node               : ', nodes[key].name)
            print('  ports amount     : ', nodes[key].ports)
            print('  interf           : ', interf)
            print('  neigh_name       : ', neigh_name)
            print('  nodes[neigh_name]: ', nodes[neigh_name])
            print('  neigh_intf       : ', neigh_intf)
            
            nodes[key].connect(pagObj, interf, nodes[neigh_name], neigh_intf)
            
            
    return nodes




def test():
    
    template = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                       'Fa0/2': {'R6': 'Fa0/2'},
                       'Fa0/3': {'R8': 'Fa0/4'}},
                'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
                'R6': {'Fa0/2': {'R4': 'Fa0/2'},
                       'Fa0/1': {'R8': 'Fa0/3'}},
                'R7': {'Fa0/1': {'R8': 'Fa0/2'}},
                'R8': {'Fa0/2': {'R7': 'Fa0/1'},
                       'Fa0/3': {'R6': 'Fa0/1'},
                       'Fa0/4': {'R4': 'Fa0/3'}}
                }
    '''
    template = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                       'Fa0/2': {'R6': 'Fa0/5'}},
                'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
                'R6': {'Fa0/5': {'R4': 'Fa0/2'}}}
    '''
    return three_tiers_topo(template)


if __name__ == '__main__':

    import win32com.client
    from win32com.client import constants as vis

    import template_duplicate_deleter as deleter
    import node_creator
    appVisio = win32com.client.gencache.EnsureDispatch( 'Visio.Application' )
    appVisio.Visible =1
    doc = appVisio.Documents.Open("C:\\Users\\jos\\Documents\\GitHub\\win32\\Netw_empty.vsdx")
    MasterFile = 'C:\\Users\\jos\\Documents\\GitHub\\win32\\nodes.vssx'    
    pagObj = doc.Pages.Item(1)
    s_test = {}
    s_test = test()
