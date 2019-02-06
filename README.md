# Geomedia-Automation
Geomedia is a Geographic Information System spatial query software. It is useful for creating visual/geometric queries on spatial objects e.g. finding out how far away two or more geometric objects are, if land parcels overlap, how close geometric assets are to the coast etc.

Uses MS COM Interop and the Hexagon/Intergraph Geomedia SDK to create a python wrapper for working with the application. This provides a base for automation projects and repetitive GIS queries. Unfortunately, the wrapper is only half-baked as I no longer have access to copy of Geomedia to develop with but some of the basic functionality is there and working. 

I created this project out of need to do automate tasks in Geomedia, but ran out of time to finish and test it. The parts of the code that have been tested are commented and often accompanied by commented out usage examples. The Geomedia SDK is already documented in the program's help files, but examples are limited to VBA and there is no wrapper or helper class for tackling big automation projects that require some level of abstraction. This creates a similar platform to script on as Esri's ArcGIS Python API.

## Creating a new workspace or opening an existing workspace is easy with the GMWrapper module:
```python
from GMWrapper import GMWrapper, GMDocument, GMServices
Geomedia = GMWrapper(Visible=False)

# Open a workspace
workspace = Geomedia.open_workspace(r"C:\my_geomedia_workspace.gws")
Geomedia.Visible = True

# Create a document and services handler
Document = GMDocument(workspace)
Services = GMServices(Document)
```

## Create a new database connection:
```python
Document.create_connection("ConnectionName", "Connection description", r"C:\path\to\some\file.mdb")
```

## Join two database tables:
```python
Document.join_tables("LEFT_DB_CONN_NAME", "LEFT_DB_TABLE", "LEFT_DB_COL_NAME", 
				 "RIGHT_DB_CONN_NAME", "RIGHT_DB_TABLE", "RIGHT_DB_COL_NAME", 
         "Inner", queryname="optional query name")
```

## Create a buffer around an object:
```python
# Defaults to metric units, because that's real engineers use in the civilised world.
Document.buffer("CONN_NAME", "TABLE_NAME", loci_distance=12, 
				 query_name="My 12m buffer", query_description="A buffer of 12 meters around the geosptial object in the TABLE_NAME",
				 merge_buffers=True)
```

## Perform and save a spatial query (looks for anything in the superset that touches anything in the filterset):
```python
Document.spatial_query("SUPERSET_CONN_NAME", "SUPERSET_TABLE", "FILTER_CONN_NAME", "FILTER_TABLE",
				   "Touches", queryname="my query name")
```
