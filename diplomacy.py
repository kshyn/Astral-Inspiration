import networkx as nx
import xlrd
from pyvis.network import Network

nt = Network('550px', '1000px')
file = "test.xls"
G = nx.Graph()
names = []
book = xlrd.open_workbook(file)
sheet = book.sheet_by_index(0)
n = 0
for row in range(sheet.nrows):
    if row == 0:
        continue
    else:
        data = sheet.row_slice(row)
        owner = data[0].value
        while n < 6:
            n = n + 1
            lord = data[n].value
            if len(lord) != 0:
               print(owner, lord)
               names.append((owner, lord))
            else:
              break
    n = 0
G.add_edges_from(names)
nt.from_nx(G)
nt.show_buttons(filter_=['physics'])
nt.save_graph('diplomacy.html')
