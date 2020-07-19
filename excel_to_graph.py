import os
import sys
import xlrd
import networkx as nx
import matplotlib.pyplot as plt

__dirname__ = os.path.dirname(os.path.realpath(__file__))


def get_graph_from(xlsx_path):
    """
    Reads an xlsx file into a networkx graph object

    Args:
        xlsx_path: Path of xlsx file to read
    
    Return:
        Networkx MultiDiGraph object
    """
    wb = xlrd.open_workbook(file_path)
    sh = wb.sheet_by_name('Sheet1')

    graph = nx.MultiDiGraph()

    for i in range(1, sh.nrows-1):
        curr_val = sh.cell_value(i, 1)
        next_val = sh.cell_value(i+1, 1)
        if (curr_val=='' or next_val==''): 
            continue
        if not graph.has_edge(curr_val, next_val):
            graph.add_edge(curr_val, next_val)

    return graph


def get_flowchart_nodes_by_type(graph):
    """
    Classifies networkx graph nodes into flowchart shape types

    Args:
        graph: A networkx graph

    Return:
        A tuple ([connector_nodes], [process_nodes], [decision_nodes])
    """
    connector = []  
    process = []
    decision = []
    for node in graph.nodes:
        if graph.out_degree(node)>1:
            decision.append(node)
        elif graph.in_degree(node)==0 or graph.out_degree(node)==0:
            connector.append(node)
        else:
            process.append(node)

    return (connector, process, decision)

## ####################
## Graph visualizations

def plot_nx_simple_graph(graph):
    """
    Draws a Simple Networkx graph

    Args:
        graph: A networkx graph
    """
    plt.figure(figsize=(15,10))
    plt.title('Simple graph using networkx')
    nx.draw_planar(flowchart, with_labels=True, node_size=1e4)
    outpath = os.path.join(__dirname__, 'networkx_simple.png')
    plt.savefig(outpath)


def plot_nx_custom_graph(graph):
    """
    Draws a Networkx graph with box shapes 
    using positions from Graphviz layout

    Args:
        graph: A networkx graph
    """
    positions = nx.nx_agraph.graphviz_layout(graph, prog='dot')
    plt.figure(figsize=(15,10))
    plt.title('Custom graph using networkx')
    plt.axis('off')
    nx.draw(graph, positions, with_labels=True, label='nx graph',
                    node_color='w', node_shape='s', node_size=1.25e3,
                    bbox=dict(facecolor='skyblue', 
                            edgecolor='b',
                            boxstyle='round,pad=1.0'))
    outpath = os.path.join(__dirname__, 'networkx_custom.png')
    plt.savefig(outpath)
    

def plot_nx_flowchart(graph):
    """
    Draws a Networkx graph with box shapes 
    using positions from Graphviz layout

    Args:
        graph: A networkx graph
    """
    circle, rect, kite = get_flowchart_nodes_by_type(flowchart)
    positions = nx.nx_agraph.graphviz_layout(graph, prog='dot')
    plt.figure(figsize=(15,25))
    plt.title('Flowchart using networkx')
    nx.draw_networkx_nodes(graph, positions, nodelist=circle, with_labels=True, 
                            node_shape='o', node_size=1.25e4)
    nx.draw_networkx_nodes(graph, positions, nodelist=rect, with_labels=True, 
                            node_shape='s', node_size=1.25e4)
    nx.draw_networkx_nodes(graph, positions, nodelist=kite, with_labels=True, 
                            node_shape='D', node_size=1.1e4)
    graph_nodes = graph.nodes()
    label_dict = dict(zip(graph_nodes, graph_nodes))    
    nx.draw_networkx_labels(graph, positions, label_dict, font_size=12)
    nx.draw_networkx_edges(graph, positions, node_size=1.5e4, width=3, arrowsize=10)
    outpath = os.path.join(__dirname__, 'networkx_flowchart.png')
    plt.savefig(outpath)


def plot_pygraphviz(graph):
    """
    Draws a networkx graph using pygraphviz 

    Args:
        graph: A networkx graph
    """
    agraph = nx.nx_agraph.to_agraph(graph)
    circle, rect, kite = get_flowchart_nodes_by_type(graph)
    for c in circle:
        agraph.get_node(c).attr['shape']='oval'
    for r in rect:
        agraph.get_node(r).attr['shape']='box'
    for k in kite:
        agraph.get_node(k).attr['shape']='diamond'
    agraph.layout(prog='fdp')
    outpath = os.path.join(__dirname__, 'pygraphviz_out.png')
    agraph.draw(outpath, format='png')


def plot_pydot(graph):
    """
    Draws a networkx graph using pydot 

    Args:
        graph: A networkx graph
    """
    pdot = nx.drawing.nx_pydot.to_pydot(graph)
    circle, rect, kite = get_flowchart_nodes_by_type(graph)
    for node in pdot.get_nodes():
        if node.get_name().strip('"') in kite:
            node.set_shape('diamond')
        elif node.get_name().strip('"') in circle:
            node.set_shape('oval')
        elif node.get_name().strip('"') in rect:
            node.set_shape('rect')
        else:
            node.set_shape('parallelogram')
    outpath = os.path.join(__dirname__, 'pydot_out.png')
    pdot.write_png(outpath)


filename = 'demo.xlsx'
file_path = os.path.join(__dirname__, filename)
flowchart = get_graph_from(file_path)

# 1. plotting network simple graph
plot_nx_simple_graph(flowchart)

# 2. plotting networkx custom shape graph
plot_nx_custom_graph(flowchart)

# 3. plotting networkx graph as flowchart
plot_nx_flowchart(flowchart)

# 4. using pygraphviz to draw 
plot_pygraphviz(flowchart)

# 5. using pydot 
plot_pydot(flowchart)


