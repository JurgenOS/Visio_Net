#-*- encoding:utf-8 -*-

import win32com.client
from win32com.client import constants as vis

class Node(object):
    '''class to create a node and drop it on the page
    name = node name
    ports = amount of ports'''

    def __init__(self, name, ports):

        self._name = name
        self._ports = ports
        self._node = None
        self._image = None
        self._connector = None
        self._links = {}
        '''
        self._links = {fa0/1: [sw2, fa0/10, link_instance]
                       fa0/2: [sw4, fa0/23, link_instance] 
                       }'''

    @property
    def name(self):
        return self._name

    @property
    def ports(self):
        return self._ports

    @property
    def image(self):
        return self._image
    @property
    def connector(self):
        #self._connector.Shapes.ItemU('StartLabel').Text = ''
        #self._connector.Shapes.ItemU('EndLabel').Text = ''
        return self._connector
    
    @property
    def links(self):
        return self._links
    
    def get_intf(self, intf):
        return self.image.Cells('Connections.X{}'.format(intf.split('/')[-1]))


    def __str__(self):
        return 'Name = {}'.format(self.name)

    def create(self, file, page,stencil, Visio, x = 0, y = 0):
        '''to create the node from Master of stencils and
        drop the node on the page on x-y coordinates
        !!!every node can have only ONE image!!!
        file = 'file_name.vssx' in the same directiory
        page = page object, that has been created in main.py script 
             like "page = doc.Pages.Item(1)"
        stencil = the name of stencil, that you can find in the stencils window
              like "L2_24_Blue"
        Visio = appVisio.Documents.Open("nodes.vssx")
        page = page object that has been created in main.py script 
             like "page = doc.Pages.Item(1)'''

        nodes = Visio.Documents.Open('{}'.format(file)) #open the Master stencils file
        self._node = nodes.Masters(stencil)
        self._connector = nodes.Masters.ItemU('MagicLink')
 
        if not self.image:
            self._image = page.Drop(self._node, x, y)
            self._image.Text = self.name
            self._image.NameU = self.name
        else:
            print('Node {} is already dropped!'.format(self._name))

    def connect(self, page, self_intf, host2, host2_intf):

        '''connect nodes between each other 
           host1 and host2 are instances of the class 'Node'
           the connector = _connector entity has been created at the line 70
        self._links = {fa0/1: [sw2, fa0/10, link_instance]
                       fa0/2: [sw4, fa0/23, link_instance] 
                       }
        sw1.connect(pagObj, 'fa0/10', sw2, 'fa0/11')
        link = pagObj.Drop(nodes.Masters.ItemU('MagicLink'), 10, 10)
        link.Shapes.ItemU('StartLabel').Text = 'sadf'
        '''        

        self._links[self_intf] = [host2, host2_intf]
        link = page.Drop(self.connector, -1, -1)
        self._links[self_intf].append(link)

        link.Cells("EndX").GlueTo(self.get_intf(self_intf))
        link.Shapes.ItemU('EndLabel').Text = '{}'.format(self_intf)

        link.Cells("BeginX").GlueTo(host2.get_intf(host2_intf))
        link.Shapes.ItemU('StartLabel').Text = '{}'.format(host2_intf)


def test():

    MasterFile = 'C:\\Users\\jos\\Documents\\GitHub\\win32\\nodes.vssx'

    sw1 = Node('SW1', 48)
    sw1.create(MasterFile, pagObj, 'L2_24_Blue', appVisio, x=5, y=5)

    sw2 = Node('SW2', 48)
    sw2.create(MasterFile, pagObj, 'L2_24_Blue', appVisio)
    
    sw1.connect(pagObj, 'fa0/10', sw2, 'fa0/11')
    sw2.connect(pagObj, 'fa0/11', sw1, 'fa0/10')
    

if __name__ == '__main__':
    test()
