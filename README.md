# Excel To Graph - Demo 
A demo script to show different ways to convert a process' steps to graph image using networkx

Uses libraries:

* networkx
* graphviz / pygraphviz / python-graphviz
* pydot

**Expected Input**: An xlsx file named 'demo.xlsx' with text values (to use as graph nodes) in column B. A sample file is included in repository.

**Script Output**: Image files (.png) for:

1. Networkx graph with planar layout (to avoid edge intersection)
2. Networkx graph with custom node shape - rectangle for all nodes
3. Networkx graph drawn with flowchart shapes
4. Pygraphviz graph with force directed placement layout
5. Pydot graph

