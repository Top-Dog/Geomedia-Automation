from win32com.client.gencache import EnsureDispatch
from win32com.client import constants as c
#from win32service import CreateService

#>>> from win32com.client.gencache import EnsureDispatch
#>>> from win32com.client import constants as c
#>>> GeoApp = EnsureDispatch("Geomedia.Application", "localhost")
#>>> GeoApp.Visible = 1
#>>> Document = GeoApp.New()
#>>>

class GMWrapper(object):
	def __init__(self, **kwargs):
		self.GeoApp = EnsureDispatch("Geomedia.Application", "localhost")
		self.DB = EnsureDispatch("Access.GDatabase")
		print "Application successfully bound!"
		self.GeoApp.Visible = kwargs.get("Visible", True)
		self.Document = self.GeoApp.Document # Contains Document objects (key is the FullName i.e path+filename)
		self.GeoApp.DefaultFilePath = r"J:\GIS\Geomedia\workspaces"
		self.GeoApp.CompanyName = "PowerNet Ltd."

		# Generate constants (doesn't seem to work if these CCIDs are created as services)
		Connection = EnsureDispatch("Geomedia.Connection")
		SpatialFilter = EnsureDispatch("GMService.ApplySpatialFilterService")
		SpatialSubPipe = EnsureDispatch("GeoMedia.SpatialSubsetPipe")
		EqiPipe = EnsureDispatch("GeoMedia.EquijoinPipe")
		BuffPipe = EnsureDispatch("Geomedia.BufferPipe")
		ConfPipe = EnsureDispatch("Geomedia.ConflictDetectionPipe")
		CoordPipe = EnsureDispatch("GMservice.CoordGeocodePipe")
		AggPipe = EnsureDispatch("Geomedia.AggregationPipe")
		ConnDialog = EnsureDispatch("GMService.NewConnectionDialog")
		ExptFC = EnsureDispatch("GeoMedia.ExportToFGDB")
		GeoStore = EnsureDispatch("Geomedia.GeometryStorageService")
		Spp = EnsureDispatch("GeoMedia.SchemaProjectionPipe") 
		Ots = EnsureDispatch("Geomedia.OutputToTableService")
		MrgPipe = EnsureDispatch("Geomedia.MergePipe")
		FxPipe = EnsureDispatch("GeoMedia.FunctionalAttribute")
		LegendEntry = EnsureDispatch("Geomedia.LegendEntry")

	def GetRGB(self, r, g, b):
		"""Sets the RGB colour as a long int, as required by the 
		FeatureSymbology object Color and FillColor properties in GeoMedia.
		This is the same as the RGB(r, g, b) function in VBA."""
		return (r & 255) + (g & 255) * 256 + (b & 256) * 256**2
		
	GetRGB = lambda self, r, g, b: r + 256 * g + b * 256**2

	def new_workspace(self, workspacename):
		# The first new document/workspace is a map window
		if self.Document is None:
			workspace = self.GeoApp.New()
			self.Document = workspace
			self.GeoApp.ActiveWindow.WindowState = c.gmwMaximize
			#self.GeoApp.ActiveWindow.NorthArrow.BackColor = self.GetRGB(200, 200, 200)
			#self.GeoApp.ActiveWindow.ScaleBar.BackColor = self.GetRGB(200, 200, 200)

			self.GeoApp.PropertySet.SetValue("Description", "GMWrapper created workspace")
			return workspace
		else:
			# We are overwriting an existing workspace, must remove it first
			print "Workspace %s already exists, please remove it first." % workspacename
			return None

	def open_workspace(self, filename):
		if self.Document is None:
			print "Opening workspace"
			workspace = self.GeoApp.Open(filename)
			print "Workspace successfully opened"
			self.Document = workspace
			self.GeoApp.ActiveWindow.WindowState = c.gmwMaximize
			#self.GeoApp.ActiveWindow.NorthArrow.BackColor = self.GetRGB(200, 200, 200)
			#self.GeoApp.ActiveWindow.ScaleBar.BackColor = self.GetRGB(200, 200, 200)

			return workspace
		else:
			return self.GeoApp.Document

	def save_workspace(self, workspacename, filename):
		if self.Document is not None :
			workspace = self.Document # Get a document (workspace object)
			try:
				workspace.Save()
			except Exception as e:
				print e
				workspace.SaveAs(filename)

	def close_workspace(self, savechnages=True, savefilename=""):
		if savechnages:
			if savefilename == "":
				self.Document.close(True)
			else:
				self.Document.close(True, savefilename)
		else:
			self.Document.close()

	@property
	def Visible(self):
		"""Get the visible state of the GeoMedia application"""
		return self.GeoApp.Visible

	@Visible.setter
	def Visible(self, state):
		"""Set the visible state of the GeoMedia application"""
		self.GeoApp.Visible = state

	def close_application(self):
		self.GeoApp.Quit()




####### GDatabase #########
class GDatabase():
	def __init__(self, filename):
		self.filename = filename
		self.DB = EnsureDispatch("Access.GDatabase") # Create a database object
		self.Tables = {} # Try and pre-populate 

	def get_table_info(self, warehouse=r"C:\warehouses\USSampleData.mdb"):
		self.DB = CreateObject("Access.GDatabase") # EnsureDispatch
		self.DB.OpenDatabase(warehouse, True, True) # OpenDatabase(name, exclusive, readonly, ModTrack, Source) - only arg1 is required
		print "DB Diagnostics:"
		print "Collating Order = " + self.DB.CollatingOrder
		print "Connection String = " + self.DB.Connect
		print "Name = " + self.DB.Name
		print "Schema updatable? = " + self.DB.SchemaUpdatable
		print "SQL conformant? = " + self.DB.SQLConformance
		print "Transactions allowed? = " + self.DB.Transactions
		print "Database updatable? = " + self.DB.Updatable
		print "Datbase version = " + self.DB.Version
		print
		print "DB table information:"
		for table in self.DB.GTableDefs: # GTableDefs() ?? # can also get tables by name e.g. objTbl = objDB.GTableDefs("tableName")
			print "\tTable name = " + table.Name
			print "\tTable created on " + table.DateCreated
			print "\tTable last updated on " + table.LastUpdated
			print "\tRecords = " + table.RecordCount
			print "\tTable updatable? = " + table.Updatable
			for field in table.GFields: # GFields() ?? # Can also get fields by name: objFld = objDB.GTableDefs(ComboBox1.Text).GFields(ComboBox2.Text)
				print "\t\tField (Column) Name = " + field.Name
				print "\t\tCollating order = " + field.CollatingOrder
				print "\t\tField data updatable? = " + field.DataUpdatable
				print "\t\tSource database = " + field.SourceDatabase
				print "\t\tSource field = " + field.SourceField
				print "\t\tSource Table = " + field.SourceTable
				print "\t\tAllow zero length? = " + field.AllowZeroLength
				print "\t\tAttributes = " + field.Attributes
				print "\t\tDefault value = " + field.DefaultValue
				print "\t\tField size = " + field.FieldSize
				print "\t\tField is required? = " + field.Required
				print "\t\tSize = " + field.Size
				print "\t\tType = " + field.Type
				if field.Type == c.gdbSpatial or field.Type == c.gdbGraphic:
					print "\t\t\tGeometric fields:"
					print "\t\t\tSubtype = " + field.SubType
					print "\t\t\tCoordinate system GUID = " + field.CoordSystemGUID
				print

	def create_db(self):
		# Create a new DB Object
		self.DB.CreateDatabase(self.filename, c.gdbLangGeneral)

	def create_db_table(self, tablename):
		table = self.DB.CreateTableDef(tablename)
		table.Name = tablename
		self.Tables[tablename] = table

	def add_table(self, tablename):
		self.DB.GTableDefs.Append(self.Tables[tablename])

	def create_db_field(self, tablename, fieldname, fieldtype, attributes, **kwargs):
		# Create the new fields (columns) in the DB
		table = self.Tables.get(tablename)
		FieldObj = table.CreateField(fieldname, fieldtype)
		FieldObj.Attributes = attributes
		FieldObj.Name = fieldname
		if kwargs.get(size):
			FieldObj.Size = kwargs.get(size, 25)
		if kwargs.get(required):
			FieldObj.Required = kwargs.get(required, False)
		if kwargs.get(defaultvalue):
			FieldObj.DefaultValue = kwargs.get(defaultvalue, "")
		if kwargs.get(allowzeroLength):
			FieldObj.AllowZeroLength = kwargs.get(allowzeroLength, True)

		if kwargs.get(geotype):
			FieldObj.Type = kwargs.get(geotype)
		if kwargs.get(subgeotype):
			FieldObj.SubType = kwargs.get(subgeotype)
		if kwargs.get(coordsystemGUID):
			FieldObj.CoordSystemGUID = kwargs.get(coordsystemGUID, True)

		table.GFields.Append(FieldObj)

	def create_db_commit(self, tablename):
		table = self.Tables.get(tablename)

	def create_db_x(self, tablename):
		table = self.Tables.get(tablename)
		objIdx = table.CreateIndex("IDidx")
		objIdxFld = objIdx.CreateField("ID")
		objIdx.GFields.Append(objIdxFld)
		objIdx.Name = "IDidx"
		objIdx.IgnoreNulls = False
		objIdx.Primary = True
		objIdx.Unique = True
		table.GIndexes.Append(objIdx)

	def create_db_meta(self, tablename):
		table = self.Tables.get(tablename)
		conn = self.create_connection("Metadata Connection", "Creates metadata for DB table", r"C:\Temp\NewDB.mdb", dbtype="Access Read Write")
		#conn = self.Connections.get(coonectionname)
		service = self.create_application_service("Geomedia.MetadataService") # test if this line works
		service.Connection = conn ##### FIX THIS LINE ##
		service.TableName = "NewTable"
		x = self.create_application_service("Geomedia.TableProperty") # test if this line works
		x.Description = "New database table"
		x.Name = "NewTable"
		x.PrimaryGeometryFieldName = "Geometry"
		x.GeometryType = 1
		service.AddTableMetadata(x)

	def create_db_add_css_row(self):
		# Adding a row (move to seperate method)
		AliasTableName = db.GAliasTable()
		rs = db.OpenRecordset("SELECT TableName From " + AliasTableName +
						" WHERE TableType='GCoordSystemTable'", c.gdbOpenSnapshot)
		CoordSystemTableName = rs.GFields.Item(0)
		rs.Close()
		db.GTableDefs.Append(table)
		db.Close()

def app_db(self):
	GDB = GDatabase(r"c:\temp\NewDB.mdb")
	GDB.create_db()
	GDB.create_db_table("NewTable")
	GDB.create_db_field("NewTable", "NumField", c.gdbDouble, c.gdbVariableField, defaultvalue=10, required=False)
	GDB.create_db_field("NewTable", "TxtField", c.gdbText, c.gdbUpdatableField, required=False, size=80, allowzerolength=False)
	GDB.create_db_field("NewTable", "ID", c.gdbLong, c.gdbAutoIncrField)
	GDB.create_db_field("NewTable", "Geometry", c.gdbSpatial, None, geotype=c.gdbSpatial, subgeotype=c.gdbLinear, required=False, coordsystemGUID=rs.GFields("CSGUID").Value)
	GDB.add_table("NewTable") # Adds the table to the DB

	# Create and define index on the "ID" Field
	#Dim objIdx As GIndex, objIdxFld As GField

	#Set objIdx = objTbl.CreateIndex("IDidx")

	#Set objIdxFld = objIdx.CreateField("ID")
	#objIdx.GFields.Append objIdxFld
	#Set objIdxFld = Nothing

	#With objIdx
	#   .Name = "IDidx"
	#   .IgnoreNulls = False
	#   .Primary = True
	#   .Unique = True
	#End With

	#objTbl.GIndexes.Append objIdx
	# self.DB.Close()

####### GDatabase #########








class GMDocument(GMWrapper):
	# Supply a GeoMedia document i.e. workspace
	def __init__(self, GeoMediaDocumentInstance):
		self.Document = GeoMediaDocumentInstance
		self.Connections = {}
		for conn in self.Document.Connections:
			self.Connections[conn.Name] = conn

	def new_map_window(self, name, legend=None):
		self.Document.NewMapWindow(legend)

	def new_data_window(self, name):
		self.Document.NewDataWindow()

	def new_layout_window(self, name):
		self.Document.CreateLayoutWindow()

	def get_application_instance(self):
		# Returns the GeoMedia application instance
		return self.Document.Parent

	def create_application_service(self, servicename):
		return self.get_application_instance().CreateService(servicename)

	def _print_rs_attributes(self, rs, *properties):
		"""Prints an array of the Names of all the attribute fields 
		(column names) in given recordset. Recordset are GField Objects.
		A GField Object represetns one column of data with a common datatype
		and set of properties. Set/Get Name and data type, geometry type (if applicable),
		value, and default value. All other properties are read-only (get properties)."""
		# print [i.Name for i in rs]
		# All properties: AllowZeroLength, Attributes, CollatingOrder, CoordSystemGUID, DataUpdatable, DefaultValue,
		# Name, Required, Zoze, SourceDatabase, SourceField, SourceTable, SubType, Type, Value
		attribs = []
		for p in properties:
			try:
				# Try a get the property
				l = [i.p for i in rs]
			except:
				# The property 'p' is not valid
				l = []
			print l
			attribs.append(l)
		return attribs

	def printcommethods(self, com_object):
		for key in dir(com_object):
			method = getattr(com_object,key)
			if str(type(method)) == "<type 'instance'>":
				print key
				for sub_method in dir(method):
					if not sub_method.startswith("_") and not "clsid" in sub_method.lower():
						print "\t"+sub_method
			else:
				print "\t",method

	####### DB (Warehouse) Connections #########
	def create_connection(self, connectionname, connectiondescription, dbfilename, dbtype="Access Read Only"):
		DatabaseTypes = {"Access Read Only" : "AccessRO.GDatabase", "Access Read Write" : "Access.GDatabase",
				   "ArcView" : "AV.GDatabase", "CAD" : "GCAD.GDatabase", "File Geodatabase Read Only" : "FGDBRO.GDatabase",
				   "File Geodatabase Read Write" : "FGDBRW.GDatabase", "GML" : "GML.GDatabase", "KML" : "KML.GDatabase", 
				   "I/CAD MAP" : "ICADMAP.GDatabase", "MapInfo" : "MI.GDatabase", "ODBC Tabular Read Only" : "ODBCRO.GDatabase",
				   "Oracle Read Only" : "OracleORO.GDatabase", "Oracle Read Write" : "OracleORW.GDatabase", 
				   "SmartStore" : "GeoMediaSmartStore.GDatabase", "SQL Server Read Only" : "SQLServerRO.GDatabase" , 
				   "SQL Server Read Write" : "SQLServerRW.GDatabase", "Text File" : "TextFile.GDatabase", 
				   "WCS" : "WCS.GDatabase", "WFS Read Only" : "WFS.GDatabase", "WFS Read Write" : "WFSRW.GDatabase", 
				   "WMS" : "WMS.GDatabase", "WMTS" : "WMTS.GDatabase "}
		if "only" in dbtype.lower():
			mode = c.gmcModeReadOnly
		else:
			mode = c.gmcModeReadWrite
		dbtype = DatabaseTypes.get(dbtype)
		assert dbtype is not None, "Not a valid database type"

		if not self.connection_exist(connectionname, dbfilename):
			ConnectionObject = self.create_application_service("Geomedia.Connection") # Geomedia seems to replace PClient in the launch string (OLE Identfier/ProgID)
			ConnectionObject.Name = connectionname
			ConnectionObject.Description = connectiondescription
			ConnectionObject.Type = dbtype # Seach Object browser for "GDatabase" to get a full list of types
			ConnectionObject.Location = dbfilename
			ConnectionObject.Mode = mode # c.gmcModeReadOnly, c.gmcModeReadWrite 

			ConnectionObject.Connect()
			
			# T: broadcast changes made by current session to all recordset listeners, F: broadcast all changes
			ConnectionObject.BroadcastDatabaseChanges(False)
		
			self.Document.Connections.Append(ConnectionObject)
			self.Connections[connectionname] = ConnectionObject
			#return ConnectionObject

	def connection_exist(self, connectionname, dbfilename):
		"""Returns true if a connection with the same name or location 
		already exisits."""
		for connection in self.Document.Connections:
			if connection.Name == connectionname or connection.Location == dbfilename:
				return True
		return False

	def disconnect(self, connectionname):
		self.Connections.get(connectionname).Disconnect()

	def connect(self, connectionname):
		self.Connections.get(connectionname).Connect()

	####### DB (Warehouse) Connections #########

	################## Pipes #################################
	# Pipes contain software components that accept some input (usually recordsets), 
	# perform a calculation, and output a new recordset. Pipes may require additional
	# paramters to filter or perform a query. Pipes are persistant throughout a Geomedia session,

	# Services are similar to pipes, but are not persistant i.e. the service is 
	# terminated when the method calling the service is completed.

	def create_originating_pipe(self, connectionname, tablename, sqlwhere=""):
		"""# Use this when recordset listeners e.g. Map views or legends need to recieve notifications of changes to the recordset"""
		#ConnectionObject = self.get_application_instance().CreateService("Geomedia.Connection")
		# Needs a connection objec, where the .Connect() method has been called
		# where pipename is type<geomedia.orginatingpipe>

		#using PClient = Intergraph.GeoMedia.PClient;
		#PClient.OriginatingPipe objOPipe = null;
		#objConn.CreateOriginatingPipe(out objOPipe);
		#objOPipe.Table = "Cities";
		#Recordset = objOPipe.OutputRecordset;
		originatingpipe = self.Connections.get(connectionname).CreateOriginatingPipe()
		originatingpipe.Table = tablename
		originatingpipe.Filter = sqlwhere
		originatingpipe.GeometryFieldName = "G3E_GEOMETRY" # Field name that contains geometry data ("Geometry")
		return originatingpipe

	def get_originating_pipes(self, connectionname):
		"""Returns a list of originating pipes on a particular connection. A
		orginating pipe defines a database query criterion based on a SQL
		where clause (Filter) and/or spatial crieria."""
		return self.Connections.get(connectionname).GetOriginatingPipes()

	def count_originating_pipes(self, connectionname):
		try:
			return self.Connections.get(connectionname).OriginatingPipeCount()
		except Exception as e:
			print e.message
			return -1
			# Could fail if the Connection.Connect() method is not called first, or
			# the connectionname is not valid

	def create_sort_pipe(self):
		"""A sort pipe orders the input recordset. Requires (1) input
		recordset and (2) collection of sort criteria."""
		return self.create_application_service("Geomedia.SortPipe")

	def create_buffer_pipe(self):
		"""Creates an output recordset that contains rows of buffers for each
		row in the input recordset. The output recordset also a query status
		field and a field to indicate the distance used to generate the buffer."""
		return self.create_application_service("Geomedia.BufferPipe")

	def create_conflict_detection_pipe(self):
		"""Pipe that detects geometry conflicts in an input recordset.
		The output recordset is the same as the input recordset except 
		that an additional column/field is added that indicates the existence
		or absence of a geometry conflict (True/False Boolean). A conflict
		between two geometries is determined by the spatial touch operator."""
		return self.create_application_service("Geomedia.ConflictDetectionPipe")

	def create_spatial_subset_pipe(self):
		return self.create_application_service("GeoMedia.SpatialSubsetPipe")

	def create_spatial_intersection_pipe(self):
		return self.create_application_service("GeoMedia.SpatialIntersectionPipe")

	def create_equijoin_pipe(self):
		return self.create_application_service("GeoMedia.EquijoinPipe")

	def create_aggregation_pipe(self):
		return self.create_application_service("Geomedia.AggregationPipe")

	def create_spp(self):
		"""Schema Projection pipe (SPP). This pipe allows the user to change
		database/recordset schema."""
		return self.create_application_service("GeoMedia.SchemaProjectionPipe")

	def create_otts(self):
		"""Output to table service (OTTS). Like a terminating pipe. Allows user to 
		create a table, append records to table, and/or update records in a table."""
		return self.create_application_service("GeoMedia.OutputToTableService")

	def create_afp(self):
		"""Creates an attribute Filter Pipe (AFP). The output recordset is filtered using a 
		SQL WHERE clause (Filter property)."""
		return self.create_application_service("Geomedia.AttributeFilterPipe")

	def create_mp(self):
		"""Create a geometry and/or attribute based merge (Merge Pipe)."""
		return self.create_application_service("Geomedia.MergePipe")

	def create_fa(self):
		"""Create a functional attribute (FA)"""
		return self.create_application_service("GeoMedia.FunctionalAttribute")

	def create_le(self):
		"""Create a legend entry"""
		return self.create_application_service("Geomedia.LegendEntry")

	def spatial_xform_pipe(self):
		return self.create_application_service("Geomedia.CSSTransformPipe")


	#def create_db(self):
	#	"""Create a Database"""
	#	return self.create_application_service("Access.GDatabase")


	################## Pipes #################################

	################# Database ###############################
	def get_data_recordset(self, connectionname):
		"""An alterntie to originating pipes."""
		# Use this add/modify/delete information in the database without notifying event listeners (application independant)
		# Similar function to createoriginatingpipe
		self.Connections.get(connectionname).Database.OpenRecordset # GDatabase.OpenRecordset 
	################# Database ###############################

	def get_query_rs(self, queryname):
		"""Returns the recordset of a saved query"""
		for q in self.Document.QueryFolder.QuerySubFolders("Queries").Queries:
			if q.Name == queryname:
				return q.Recordset
		return None


	def save_query(self, recordset, queryname, querydescription=""):
		"""Save a query to the queries folder in the Geomedia workspace.
		B/c GetExtension is being used, a OriginatingPipe pipe ojects must 
		be used to generate the recordset."""
		query = recordset.GetExtension("Name")
		query.Name = queryname
		query.Description = querydescription
	
		# Make the query availiable via the Tools->Queries menu in the GUI
		try:
			self.Document.QueryFolder.QuerySubFolders("Queries").Queries.Append(query)
		except Exception as e:
			if e.message == "A query by this name already exists.  Query names must be unique.  Please enter a new query name.":
				print 'The query name "%s" already exists.' % queryname
			else:
				print 'Could not save query "%s"' % queryname
				print e.message


	####### Sptial Query #######################
	def _swap_recordsets(self, rs1, rs2):
		"""Swap recordsets. Example usage: rs2, rs1 = swap(rs1, rs2)"""
		#temp = Recordset1
		#Recordset1 = Recordset2
		#Recordset2 = temp
		return rs2, rs1

	# Spatial query pipe is depricated, so use spatial subset pipe and filter on originating pipe
	def spatial_query(self, supersetconnectionname, supersettable, filterconnection, filtertable, 
				   querytype, **kwargs):
		# 1. Dimension new variables (if required).

		# 2. Create a recordset containing feature data to be filtered. A subset of this recordset is returned.
		if supersetconnectionname == "Queries":
			Recordset1 = self.get_query_rs(supersettable)
		else:
			pipe1 = self.create_originating_pipe(supersetconnectionname, supersettable, kwargs.get("inputfilter", "")) 
			Recordset1 = pipe1.OutputRecordset

		# 3. Create a recordset containing feature data that will be filtering the input data. These geometries are used as filtering criteria in this spatial query.
		if filterconnection == "Queries":
			Recordset2 = self.get_query_rs(filtertable)
		else:
			pipe2 = self.create_originating_pipe(filterconnection, filtertable)
			Recordset2 = pipe2.OutputRecordset
		# Only need to do this if pipe if shared between recordset 1 and 2
		# System.Runtime.InteropServices.Marshal.ReleaseComObject(Pipe)

		# 4. Ensure that both recordsets are using a common projection (the map view projection).
		Recordset1 = self.recordset_xform(Recordset1) # Transform NZMG projection to NZTM
		Recordset2 = self.recordset_xform(Recordset2) # Transform NZMG projection to NZTM 

		# 5. Create the spatial query itself (NB: the recordset with the filter is the 2nd paramter).
		#gmsqMeet 1 Includes only features referenced by the first record set that meet features referenced in the second record set.
		#gmsqOverlap 2 Includes only features referenced by the first record set that overlap features referenced in the second record set. 
		#gmsqContains 3 Features referenced by the first record set contain features referenced in the second record set. 
		#gmsqContainedBy 4 Features referenced by the first record set are contained in the second record set. 
		#gmsqEntirelyContains 5 Features referenced by the first record set completely contain the features referenced in the second record set. The boundaries of the features cannot touch in any way; everything must be interior. 
		#gmsqEntirelyContainedBy 6 Features referenced by the first record set are completely contained in the second record set. The boundaries of the features cannot touch in any way; everything must be interior. 
		#gmsqSpatiallyEqual 7 Includes only features referenced by the first record set that are spatially equal to features referenced in the second record set. 
		#gmsqTouches 8 Includes only features referenced by the first record set that intersect features referenced in the second record set. 
		#gmsqWithinDistance 9 Includes only features within a specified distance (set by the Distance property) in the output record set. Equal to gmsqTouches plus the defined distance. (Within distance not availiable for intersection pipe?)
		QueryTypes = {"Meet" : c.gmsqMeet,  "Overlap": c.gmsqOverlap, "Contains" : c.gmsqContains,
				"Contained By" : c.gmsqContainedBy, "Entirely Contains" : c.gmsqEntirelyContains,
				"Entirely Contained By" : c.gmsqEntirelyContainedBy, "Spatially Equal" : c.gmsqSpatiallyEqual,
				"Touches" : c.gmsqTouches, "Within Distance" : c.gmsqWithinDistance}
		querytype = QueryTypes.get(querytype)
		assert querytype is not None, "Not a valid query type"

		# Get the names of geometry fields for each of the input recordsets
		inputgeoname = self.get_geometry_field_name(Recordset1)
		filtergeoname = self.get_geometry_field_name(Recordset2)
		
		# Perform a move on the recordset to reset the cache (required to increase performance of spatial pipes)
		# Need to do this or not access the geometry of the recordset beforehand.
		if not (Recordset1.BOF and Recordset1.EOF):
			Recordset1.MoveLast()
			Recordset1.MoveFirst()
		if not (Recordset2.BOF and Recordset2.EOF):
			Recordset2.MoveLast()
			Recordset2.MoveFirst()

		# https://hgdsupport.hexagongeospatial.com/API/GeoMedia/Building%20on%20the%20GeoMedia%20Engine/#query_3.htm (also in help file)
		SpatialQuery = self.create_spatial_subset_pipe()
		SpatialQuery.NotOperator = kwargs.get("querynegation", False)
		SpatialQuery.InputGeometryFieldName = inputgeoname
		SpatialQuery.InputRecordset = Recordset1 # Specifies the recordset to be filtered (Left - InputRecordset)
		SpatialQuery.FilterGeometryFieldName = filtergeoname
		SpatialQuery.FilterRecordset = Recordset2 # Recordset whose geometries are to be used for filtering the input recordset (Right - FilterRecordset)
		SpatialQuery.OutputStatusFieldName = "STATUS" # The field name in the output recordset that contains the status of the spatial query (contains text if the query failed)
		SpatialQuery.SpatialOperator = querytype
		SpatialQuery.Distance = kwargs.get("distance", 0) # Only relevant if the c.WithinDistance operator is used
		RecordsetOut = SpatialQuery.OutputRecordset # Contains all the orginal fields contained in the input recordset as well as the 'status field' which contains text in the event a failure to determine the spatial relationship

		# 6. Create the Query object and make it availiable to the workspace
		#print "Input Recordset sizes: %d, %d" % (Recordset1.RecordCount, Recordset2.RecordCount)
		#print "Output Recordset size: %d" % (RecordsetOut.RecordCount)
		if kwargs.get("queryname"):
			self.save_query(RecordsetOut, kwargs.get("queryname"), kwargs.get("querydescription", ""))
		
		# Generate the recordset and return it
		return RecordsetOut()


	def get_geometry_field_name(self, recordset):
		"""Return the geometry field name of the first geometry field in a recordset,
		if there is one. See Field Type Constants for other types including gdbGraphic,
	    gdbLongBinary, gdbGUID  etc."""
		rs = recordset()
		for field in rs:
			if field.Type == c.gdbSpatial:
				return field.Name
		return ""

	####### Sptial Query #######################

	####### Join Query #######################
	def join_tables(self, leftconnection, lefttable, leftcolname, rightconnection, righttable, rightcolname, jointype, **kwargs):
		# Make sure the join is valid
		JoinTypes = {"Inner" : c.gmejpInner, "Right Outer" : c.gmejpRightOuter,
			  "Left Outer" : c.gmejpLeftOuter, "Full Outer" : c.gmejpFullOuter,
			  "Union" : c.gmejpUnion}
		jointype = JoinTypes.get(jointype)
		assert jointype is not None, "Not a recognised join type (%s)" % jointype

		# Set the left and right tables to join
		leftpipe = self.create_originating_pipe(leftconnection, lefttable)
		rightpipe = self.create_originating_pipe(rightconnection, righttable)

		# Set the output table
		MyEquijoinPipe = self.create_equijoin_pipe()
		MyEquijoinPipe.LeftRecordset = leftpipe.OutputRecordset
		MyEquijoinPipe.RightRecordset = rightpipe.OutputRecordset
		MyEquijoinPipe.JoinFieldNames = [[leftcolname], [rightcolname]]
		MyEquijoinPipe.JoinType = jointype
		RecordsetOut = MyEquijoinPipe.OutputRecordset

		# Store a query
		if kwargs.get("queryname"):
			self.save_query(RecordsetOut, kwargs.get("queryname"), kwargs.get("querydescription", ""))
		
		# Generate the recordset and return it
		return RecordsetOut()
	
	def remove_query(self, queryname, subfoldername="Queries"):
		QueryFolder = self.Document.QueryFolder.QuerySubFolders(subfoldername)
		QueryFolder.Queries(queryname).Remove() # QueryFolder.Queries.Remove(querynumber)

		#Dim QryFld As QueryFolder
		#QryFld = Application.Document.QueryFolder.QuerySubFolders("MyFolder")
		#QryFld.Queries("MyQuery").Description = "A new description"
		#QryFld = Nothing


	####### Join Query #######################

	####### Buffer Query #######################

	def buffer(self, connectionname, tablename, **kwargs):
		# Get optional specfiers
		distancetype = kwargs.get("distance_type", "Constant")
		edgetype = kwargs.get("edge_type", "Round")
		locidistance = kwargs.get("loci_distance", 1)
		distanceunit = kwargs.get("distance_unit", "m") # Default to meters (m). Real engineers use metric.

		EdgeType = {"Round" : c.gmbztLinearRoundEnd , "Square" : c.gmbztLinearSquareEnd}
		edgetype = EdgeType.get(edgetype)
		assert edgetype is not None, "Not a valid edge type"

		DistanceType = {"Constant" : c.gmbzConstantDistance , "Variable" : c.gmbzVariableDistance}
		distancetype = DistanceType.get(distancetype)
		assert distancetype is not None, "Not a valid distance type"

		OutPipe = self.create_originating_pipe(connectionname, tablename)

		# using PPipe = Intergraph.GeoMedia.PPipe;
		# objBP = new PPipe.BufferPipe();
		BufferPipe = self.create_buffer_pipe()
		BufferPipe.BufferType = edgetype
		BufferPipe.InputRecordset = OutPipe.OutputRecordset
		BufferPipe.InputGeometryFieldName = "G3E_GEOMETRY" # Set the field name in the input recordset which contains the geometry to be buffered ("Geometry")
		BufferPipe.OutputGeometryFieldName = "BufferGeometry" # Set the field name in the output recordset which contains the buffered geometry

		# objUOM = new PCSS.UnitsOfMeasure();
		# objBP.InputDistanceUnit = objUOM.GetUnitID(PCSS.UnitTypeConstants.igUnitDistance, "mi", null);
		# using PCSS = Intergraph.GeoMedia.PCSS;
		# ... seems that PCSS is the same as the GeoApp object
		# PCSS.UnitTypeConstants.igUnitDistance (replace with c.igUnitDistance)
		UOM = self.get_application_instance().UnitsOfMeasure
		
		# set to meters, what real engineers use
		BufferPipe.InputDistanceUnit = UOM.GetUnitID(1, distanceunit, True) # 1 = igUnitDistance
		#The value in this property must be the unit ID for any distance unit 
		#of measure (unit type = igUnitDistance) defined by the UnitsOfMeasure object (UOM). 
		#Default value is igDistanceMeter.
		if distancetype == DistanceType.get("Constant"):
			# PPipe.GMBufferZoneDistanceConstants.gmbzConstantDistance
			BufferPipe.DistanceType = c.gmbzConstantDistance
			BufferPipe.InputDistance = str(locidistance) # "0:20"
		else:
			# The name of the field within the input recordset that contains 
			# the distance for the specific geometry to be buffered.
			# A NULL (None) value results in no buffer zone being generated
			BufferPipe.DistanceType = c.gmbzVariableDistance
			BufferPipe.InputDistanceFieldName = "VariableDistanceToBeBuffered"
		RecordsetOut = BufferPipe.OutputRecordset

		# Merge buffer zones if requested
		if kwargs.get("merge_buffers", False):
			if not (RecordsetOut.BOF and RecordsetOut.EOF):
				# Perform a move on the recordset to reset the cache (requierd to increase perfomace of spatial pipes)
				RecordsetOut.MoveLast()
				RecordsetOut.MoveFirst()
			# Add a functional attribute that merges the geometries
			RecordsetOut = self.merge_recordset(RecordsetOut, functionalmergefields=
				[(BufferPipe.OutputGeometryFieldName, "Merge(%s)" % BufferPipe.OutputGeometryFieldName)])

		# Store a query
		if kwargs.get("queryname"):
			self.save_query(RecordsetOut, kwargs.get("queryname"), kwargs.get("querydescription", ""))
		
		# Generate the recordset and return it
		return RecordsetOut()

	####### Buffer Query #######################

	def conflict_detection(self):
		# Conflict detection (i.e. where the buffer geometries touch)
		#ConflictPipe = self.create_conflict_detection_pipe()
		#ConflictPipe.InputRecordset = BufferPipe.OutputRecordset # Fields: buffer area objects, query status field, distance value of buffer
		#ConflictPipe.InputGeometryFieldName = BufferPipe.OutputGeometryFieldName
		#ConflictPipe.OutputConflictFieldName = "Conflict"
		pass

	def functional_attribute(self, expression, outputfieldname):
		"""Create a functional attribute handler for one field (output column)."""
		FA = self.create_fa()
		FA.Expression = expression 
		FA.FieldName = outputfieldname # set the name of the output field i.e. the field (column) where the functional attribute is to be stored
		return FA

	def merge_recordset(self, inputrecordset, functionalmergefields=[("BufferGeometry", "Merge(BufferGeometry)")], mergefields=[]):
		# include only the name of the appropriate geometry field in the MergeFieldNames property. (features that touch are merged, only 1 geometry field can be merged) 
		# To perform attribute-based merging, 
		# include only tabular field(s) in the MergeFieldNames property. 
		# To perform combination merging, 
		# include both a geometry field and a tabular field(s) in the MergeFieldNames property. (attributes that are identical are merged)
		# NULL values merge with nothing (not even other NULLs)
		MP = self.create_mp()
		MP.InputRecordset = inputrecordset
		# Sets the names of the fields used to determine which features to
		# merge, based on the equivalence of the attribute values in tabular fields
		# and/or the spatial coincidence ("touch" operator) of geometry values in geometry fields.
		fields = []
		for field in functionalmergefields:
			fields.append(field[0])
		MP.MergeFieldNames = fields + mergefields
		# At least one functional attribute is required to outputted
		for functionalattribute in functionalmergefields:
			MP.OutputFunctionalAttributes.Append(
				self.functional_attribute(functionalattribute[1], functionalattribute[0]))
		return MP.OutputRecordset
	
	###################### Transfomations #####################
	# Trasform the recorset projection to the map view projection
	def transform_projection(self, connectionname, tablename):
		MapView.CoordSystemMgr.CoordSystem.BaseStorageType = c.csbsProjected
		MapView.CoordSystemMgr.CoordSystem.RefSpaceMgr.ProjSpace.ProjAlgorithmVal = c.cspaUniversalTransverseMercator

		opipe = self.create_orginatingpipe(connectionname, tablename, CoordSystemMgr=MapView.CoordSystemMgr)

		Xformpipe = self.spatial_xform_pipe() #self.create_service("Geomedia.CSSTransformPipe")
		Xformpipe.InputRecordset = opipe.OutputRecordset
		Xformpipe.InputGeometryFieldName = "G3E_GEOMETRY"
		Xformpipe.CoordSystemMgr = MapView.CoordSystemMgr # Define the coord system of the outputrecordset
		return Xformpipe

	def recordset_xform(self, rs):
		"""This method is currently being used"""
		Xformpipe = self.spatial_xform_pipe()
		Xformpipe.InputRecordset = rs
		Xformpipe.InputGeometryFieldName = self.get_geometry_field_name(rs)
		Xformpipe.CoordSystemsMgr = self.Document.Parent.Windows(1).MapView.CoordSystemsMgr # The first window is always the map view by default
		return Xformpipe.OutputRecordset


	###################### Legend #############################

	nNextColor = 0
	def _GetStyleObject(self, iGeometryType):
		if iGeometryType == c.gdbAreal:
			# should return objstyle
			# Dim objstyleService As New PView.StyleService
							
			objstyleService.GetStyle("Area Style", objstyle)
			# OSS = xxx.GetStyle("Area Style")
			# OSS.StyleDefinitions(0).StyleDefinitions(0).StyleProperties(c.gmgroPropertyUnitType).Value = Pc.gmgroUnitTypeScalingPaper
			# OSS.StyleDefinitions(0).StyleDefinitions(0).StyleProperties(c.gmgroPropertyColor).Value = QBColor(nNextColor)
			# OSS.StyleDefinitions(0).StyleDefinitions(0).StyleProperties(c.gmgroPropertyWidth).Value = 50
								
		elif iGeometryType == c.gdbPoint:
			objStyleService.GetStyle("Font Style", objstyle)
			# OSS = xxx.GetStyle("Font Style")
			# OSS.StyleProperties(c.gmgroPropertyColor).Value = QBColor(nNextColor)
			# OSS.StyleProperties(c.gmgroPropertyUnitType).Value = gmgroUnitTypeScalingPaper
			# OSS.StyleProperties(c.gmgroPropertyFontName).Value = "Wingdings"
			# OSS.StyleProperties(c.gmgroPropertyCharacterString).Value = 74
			## OSS.StyleProperties(c.gmgroPropertySize) = 200
		elif iGeometryType == c.gdbLinear:
			objStyleService.GetStyle("Simple Line Style", objstyle)
			# OSS = xxx.GetStyle("Simple Line Style")
			# OSS.StyleProperties(c.gmgroPropertyUnitType).Value = c.gmgroUnitTypeScalingPaper
			# OSS.StyleProperties(c.gmgroPropertyColor).Value = QBColor(nNextColor)
			# OSS.StyleProperties(c.gmgroPropertyWidth).Value = 400
				
		nNextColor += 1
		if nNextColor > 14:
			nNextColor = 0	
		return OSS


	def create_legend_entry(self, recordset):
		LE = self.GeoApp.CreateService("Geomedia.LegendEntry")
		LE.InputRecordset = recordset
		LE.Style = self._GetStyleObject(c.gdbAreal)
		LE.GeometryFieldName = "G3E_GEOMETRY"
		# Dim objproperty As New PDBPipe.GMProperty (has properties Name, Value)
		# LegendEntry Object
		#LE.PropertySet.Append(objproperty)
	
		# Create the legend itself
		MapView = self.GeoApp.Document.GMMapView1
		Legend = self.GeoApp.CreateService("Geomedia.Legend")
		MapView.Legend = Legend
		GMLegendView1.Legend = MapView.Legend # somehow create GMLegendView1 object
		# GMLegendView1 has Top, Left, Width, Height properties that might need to be set too
	
		# Set the coord system of the map view to the same as the recordset's warehouse source
		CoordSysMgr = self.GeoApp.CreateService("CoordSystemsMgr")
		CoordSysMgr.CoordSystem.GUID = recordset.GFields("G3E_GEOMETRY").CoordSystemGUID
		MapView.CoordSystemsMgr = CoordSysMgr
	
		if Legend.LegendEntries.Count() == 0:
			Legend.LegendEntries.Append(LE)
		else:
			Legend.LegendEntries.Append(LE, 1)
		LE.LoadData()
		MapView.Fit()
		MapView.Refresh()

		#LE = create_servie("Geomedia.LegendEntry")
		#LE.GeometryFieldName = "G3E_GEOMETRY"
		#LE.InputRecordset = Xformpipe.OutputRecordset
		#LE.Style = GetStyle("Area Style") # PView.StyleService.GetStyle
		#LE.Locatable = True

		##PDBP (Database) pipe
		#property1 = PDBPipe.GMProperty
		#property1.Name = "Title"
		#property1.Value = ""

		#property2 = PDBPipe.GMProperty
		#property2.Name = "Subtitle"
		#property2.Value = ""

		#LE.PropertySet.Append(property1) 
		#LE.PropertySet.Append(property2)


	# Document Coord System
	def temp_legend_mapview(self):
		# GMMapView1.CoordSystemsMgr = new PCSS.CoordSystemsMgr();
		self.Document.CoordSystemsMgr
		
		self.get_application_instance().ActiveWindow.MapView.Legend
		rs = legend.LegendEntries(1).InputRecordset
		opipe = rs.GetExtension("OriginatingPipe")
		self.get_application_instance().SetStatusBar("..." + opipe.Table, 3)


class GMServices(GMDocument):
	"""Services are Geomedia software components that accept some user input,
	perform calculations, and output a set of results.
	ApplySpatialFilterService, LayoutWindowPlacementService, and
	RenderToRasterFileService  do not apply to GeoMedia objects."""
	def __init__(self, GeoMediaDocumentInstance):
		self.DocumentHandler = GeoMediaDocumentInstance
		self.Document = self.DocumentHandler.Document
		self.GeoApp = self.Document.Parent
		self.Connections = GeoMediaDocumentInstance.Connections

	def show_warehouse_connections_dialog(self, warehousefilelocation=""):
		"""Show the GeoMedia document connections dialog. Returns 1 for 'OK' or 
		2 for 'Cancel'."""
		ConnectionsDialog = self.GeoApp.CreateService("GMService.NewConnectionDialog")
		ConnectionsDialog.Connections = self.Document.Connections # Load the existing connections into the dialog
		ConnectionsDialog.AdvancedFeatureModelVisible = True # Show the enable advanced feature model checkbox
		ConnectionsDialog.AuxiliaryPath = warehousefilelocation # Set/get the default warehouse location
		ConnectionsDialog.Caption = "New Connection" # Caption/title for the dialog box
		ConnectionsDialog.PreferenceSet = self.GeoApp.PreferenceSet.GetValue("WarehouseFileLocation", warehousefilelocation, c.gmpUser)
		return ConnectionsDialog.Show()

	def _get_rs(self, **kwargs):
		rs = None
		# Get RS from connection
		if kwargs.get("connectionname") and kwargs.get("connectionname"):
			SrcOPipe = self.DocumentHandler.create_originating_pipe(kwargs.get("connectionname"), 
										   kwargs.get("tablename"))
			rs = SrcOPipe.OutputRecordset
		# Get RS from query
		elif kwargs.get("queryname"):
			for query in self.Document.QueryFolder.QuerySubFolders("Queries").Queries:
				if kwargs.get("queryname") == query.Name:
					rs = query.Recordset
		# .GetExtension("OriginatingPipe")
		return rs

	def table_service(self, inputrs, destconnectionname, desttablename, tablemode, attributestoinclude=[], schemamode="Exclude"):
		"""Create a new table in an existing database. Or append to an old one.
		@param attributestoinclude: the column names/fields/attributes to copy from the source to the destination table."""
		modes = {"Create New Table" : c.gmopmCreateNewTable, "Append Table" : c.gmopmAppendToTable,
			"New or Append" : c.gmopmNewOrAppend, "Update" : c.gmopmUpdateTable,
			"Append and Update" : c.gmopmAppendAndUpdateTable, "Force Append" : c.gmopmForceAppendToTable}
		rs = inputrs
		if attributestoinclude:
			# Change the input table schema before writing it to warehouse
			rs = self.change_table_schema(SrcOPipe.OutputRecordset, attributestoinclude, schemamode)

		DestConn = self.Connections.get(destconnectionname)
		OTTS = self.create_otts() # Terminating service/pipe OutputToTableService Object (output to table service)
		OTTS.DisableModificationLogging = True # Improves performance by turning off modification logging for data servers that support it
		OTTS.InputRecordset = rs #SrcOPipe.OutputRecordset (use this if not attribute filtering is needed)
		OTTS.OutputTableName = desttablename
		OTTS.OutputMode = modes.get(tablemode, c.gmopmNewOrAppend) # Create a new db table (not exist), or append (exist)
		OTTS.NewTableKeyMode = c.gmntkmNewKey # A new autonumber field is added to the db table and is desiganted the new key
		OTTS.NewTableAutonumberMode = c.gmntamPreserveValues # Preserve field data but change it to gdbLong
		OTTS.OutputConnection = DestConn
		# Only needed if the table has to be created for the first time (ignored if a value already exists)
		#OTTS.OutputCoordSystem = self.Document.CoordSystemsMgr.CoordSystem # self.GeoApp.Windows(1).MapView.CoordSystemsMgr.CoordSystem
		OTTS.OutputLogFileName = "C:\temp\GeoMediaTableService.log"
		OTTS.Execute()

		# Input recordset's coordinate system is missing. [no coord system supplied]
		# Unable to get the coordinate system of the input recordset primary geometry [NZTM supplied]

		# e.g. rs.GetExtension("ExtendedPropertySet").GetValue("PrimaryGeometryFieldName") [GRecordset type]
		# t = rs.GFields("BufferGeometry") [GField type]

		# Let other service listeners know that something has changed in the DB
		DestConn.BroadcastDatabaseChanges()

	def change_table_schema(self, inputrecordset, fields, mode, readonly=False):
		"""Returns a subset of the original recordset with either the spesfied 
		fields either removed or only including the specsified fields."""
		modes = {"Include" : c.gmsppIncludeFields, "Exclude" : c.gmsppExcludeFields}
		SPP = self.create_spp() # Schema projection pipe
		#SPP.GetExtension("ExtendedPropertySet") # Stores misc. user data, 1 collection unit (Property name, value, status). Access as: <Object>.PropertySet
		#SPP.FieldDefinitions 
		SPP.FieldList = fields
		SPP.FieldListMode = modes.get(mode, c.gmsppIncludeFields)
		SPP.ForeceReadOnly = readonly
		SPP.InputRecordset = inputrecordset
		return SPP.OutputRecordset

	def rename_field(self, inpipe, oldname, newname):
		"""Rename a DB field.
		@inpipe: an orginating pipe object"""
		SPP = self.create_spp() # Schema projection pipe
		SPP.InputRecordset = inpipe.OutputRecordset
		SPP.FieldListMode = c.gmsppExcludeFields
		SPP.FieldList = []

		# Rename the POP attribute to POPULATION and make it read-only (sets the DataUpdatable property of the output field to False)
		FiledDefinition = SPP.FieldDefinitions.Add(oldname) # (optional) The collection of field definitions that override the input field metadata definitions.
		FiledDefinition.OutputName = newname
		#FiledDefinition.ForceReadOnly = True
		return SPP
		
	def export_feature_class_GEODB(self, connectionname, outtablename, append=False):
		# File Geo Database (FGDB) -- THIS IS NOT FOR EXPORTING TO A FEATURE CLASS!!!!
		"""Export feature class geometry and/or attributes from any warehouse connection to
		geodatabase file format. User can either append to an existing geodatabase file or 
		create a new one."""
		Conn = self.DocumentHandler.create_connection(connectionname, 
			"", dbfilename)
		OPipe = self.DocumentHandler.create_originating_pipe(connectionname, outtablename)
		FGDBObject = self.DocumentHandler.create_application_service("GeoMedia.ExportToFGDB")
		FGDBObject.InputRecordset = OPipe.OutputRecordset
		FGDBObject.Append = append
		FGDBObject.OutputFolderName = "Output Folder name"
		FGDBObject.OutputFeatureClassName = "Output table name"
		FGDBObject.OutputGeometryDimension = 2 # 2 or 3, 3 to export 3D
		FGDBObject.OutputGeometrySubtype = c.gdbLinear #gdbPoint, gdbLinear, or gdbAreal (must be set when Feature Class geometry field type is gdbAnySpatial)
		# Spatial reference system (NZGD49)
		FGDBObject.TargetEPSGCode = "NZGD2000" # Optional? NZGD49
		FGDBObject.Execute()

	
	def create_new_db(self, connection, dbname):
		pass

	def spatial_filter(self, filtertype):
		# still need to find out how to select a geometry object (you can't create them according to the docs)
		"""This is not the same as a spatial query! This is a spatial filter to limit
		how much infomation is displayed on the mapview (fence)."""
		# https://support.hexagonsafetyinfrastructure.com/infocenter/index?page=content&id=CE2027

		QueryTypes = {"Entirely Inside" : c.gdbEntirelyContains,
				"Overlap" : c.gdbTouches, "Coarse Overlap" : c.gdbIndexIntersect,
				"Inside" : c.gdbContains}

		objASFS = self.GeoApp.CreateService("GMService.ApplySpatialFilterService")
		objGSS = self.GeoApp.CreateService("Geomedia.GeometryStorageService") # Geomedia instead of PClient
		
		# Create a binary GDO geometry blob to store geometric infomation in a databse GDO GField (type gdbSpatial)
		# GeometryStorageService.GeometryToStorage(spatial filter geometry object, binary blob)
		# Converts displayed GIS data to binary data in a table (StorageToGeometry is the oppisite)
		objGSS.GeometryToStorage(objSFGeom, blob) # where blob is just an 'object' and objSFGeom is a spatial filter geometry
		# in python this method probably returns a blob object and takes only 1 parameter
		
		objASFS.Application = self.GeoApp 
		objASFS.SpatialFilterOperator = QueryTypes.get(filtertype)
		objASFS.SpatialFilterGeometry = blob
		objASFS.SpatialFilterName = "SpatialFilter1"
		objASFS.UseGeometryMBR = False # T: MBR of Geometry, F: Geometry as is
		objASFS.Execute()

		# self.Document.QueryFolders ("SpatialModels").QuerySubFolders.Queries 

	def get_geometry_object(self, name):
		# self.GeoApp.ActiveWindow.MapView.Legend
		# Mapview.HighlightedObjects.Item(2).GeometryFieldName
		# self.Document.Parent.Windows(1).Type # Type is either MapWindpw, LayoutWindow, DataWindow etc.
		# self.Document.Parent.Windows(1).Caption
		pass

class GMEvents(GMDocument):
	def __init__(self, GMDocumentInstance):
		self.Document = GMDocumentInstance
		self.GeoApp = self.Document.Parent


# How GetExtension("xxx") works... works on GDatabase object
# See GDO Automation >> Properties for more infomation
# This method seems to be able to get GeoMedia COM objects from existing objects,
# instead of having to build them up from basic onnection objects. E.g. get an orginating pipe
# without building a connection object:
# GeoApp = GetObject(, "Geomedia.Application")
# objMV = GeoApp.ActiveWindow.MapView
# objRS = objMV.Legend.LegendEntries(1).InputRecordset
# objOPipe = objRS.GetExtension("OriginatingPipe")

## Examples:
# Name: For the GDatabase object, the Name property returns a string specifying the database. For some this is the path to a file. For others it is an identifier.
# For objects other than the GDatabase object, the Name property returns the name of the object. 

# These are the valid strings for Extension Name the call "Recordset.GetExtension(Extension Name)":
# (from GRexordsets)
# Name
# Notification
# ExtendedPropertySet		This is the only common GDO extension across all objects
# OriginatingPipe
# SpatialQueryPipe
# AttributeFilterPipe
# CSSTransformPipe
# SortPipe
# CenterPointPipe
# EquijoinPipe
# GraphicsTextPipe
# MovePipe
# SchemaProjectPipe
# SpatialFilterPipe
# SpatialDifferencePipe
# AddressGeocodePipe
# CoordGeocodePipe
# Errors

# Help search term: Extensions & Extended Properties 

# GetExtension, which is implemented on the GDatabase, GTableDef, GIndex, GRecordset, and GField objects, 
# is the method by which a data server or pipe provides access to objects that add functionality outside 
# the scope of GDO. The GDO specification levies no requirements on the returned extension object except 
# that it must be an OLE Automation interface.

# The syntax for retrieving the extension object is:
# dim objExt as Object ' Can also be strongly typed
# set objExt = <GDO Object>.GetExtension("ExtensionName")

# The extensions and extended properties in any given situation are only made known via th edocumenttaion.

# Extended properties are stored a name:value pair. They belong to the ExtendedPropertySet  collection, but 
# apply to the parent GDO object. The syntax for retrieiving an extended property:
# dim strValue as Variant
# strValue = <ExtendedPropertySet object retrieved through GetExtension>.GetValue("Property")

# op=self.create_originating_pipe("PNLPROD_DDC", "V_CRS_ROADCNTR_1")
# op.OutputRecordset.GetExtension("ExtendedPropertySet").GetValue("PrimaryGeometryFieldName")