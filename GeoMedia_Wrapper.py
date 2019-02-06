from GMWrapper import GMWrapper, GMDocument, GMServices


if __name__ == "__main__":
	Geomedia = GMWrapper(Visible=False)

	# Open the 2016 workspace
	workspace = Geomedia.open_workspace(r"I:\mygeosspace.gws")
	Geomedia.Visible = True

	# Create a document handler
	Document = GMDocument(workspace)
	Services = GMServices(Document)

	# Create a connection to the 2017 DB
	#Document.create_connection("Info2017", "Information 2017",
	#						 r"I:\2017\Info2017Geomedia.mdb")

	# Join the 2017 EHV Cable Table
	#Document.join_tables("PROD_DDC", "V_EHVCOND_LN", "G3E_FID", 
	#			 "Info2017", "CABLES2017_04_01", "G3E_FID", "Inner", queryname="Test join qry...")

	# Create a buffer around road centre lines
	Document.buffer("PROD_DDC", "V_CRS_ROADCNTR_1", loci_distance=12, 
				 query_name="My road buffer merge(1)", query_description="my road buff desc.",
				 merge_buffers=True)

	# Perform a spatial query
	Document.spatial_query("Queries", "2016 LV", "Area", "Information_Urban_Areas",
				   "Touches", queryname="1 2016 LV Urban")

	# Perform an attribute selection query
	#Document.


	# Output feature class to database
	#rs = Services._get_rs(queryname="My road buffer merge(1)")
	#Services.table_service(rs, 
	#					"Info2014", "MyRoadBuffs", 
	#					"New or Append")
	# We need to output to a GeoMedia DB, not a normal Access DB.



#from win32com.client.gencache import EnsureDispatch, Rebuild
#from win32com.client import constants as c
#GeoApp = EnsureDispatch("Geomedia.Application", "localhost")
#Document = GeoApp.Open(r"I:\Information2016.gws")
#GeoApp.Visible = 1
#Rebuild()
#from win32com.client import makepy
#import sys
#sys.argv = ["makepy", r"C:\Program Files (x86)\Hexagon\GeoMedia Professional\Program\GeoMedia.exe"]
#makepy.main()
#c.gmejpInner



#ConnObj = GeoApp.CreateService("Geomedia.Connection")
#ConnObj.Name = "My new connection"
#ConnObj.Location = r"I:\Info2017Geomedia.mdb"
#ConnObj.Connect()

