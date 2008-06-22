Steve Tessler
2008-06-02
StreamStats DB Update Final notes

These notes describe some issues discovered with StreamStats data from 2008-05-28 (via KRies).
Those data were loaded into the new structure mostly intact.
The items listed below indicate issues encountered and describe any data not loaded.
The 'not loaded' records are provided in a separate Access mdb file.
A few other issues are mentioned at the bottom of these notes.
Other suggested 'future' changes are indicated in the Excel change file.

Recent discussions re: NHD are not incorporated.
See notes below and in the Excel change doc about 5 new Station fields.
If the values will be stored in an NHD table they should be removed from Station.

---------------------------------

Loading Problems

---------

1.  Station table - the most recent DB had 29,138 records

 a  StaID 10040000 did not load from current dataset.
    Problem determined to be zero length string block on new field.
    ZLS in 'Directions' field was cleared and record loaded.

 b  QA Note - One (1) StaID in Station is not associated with any Statistics
      01482900
      
    Station table summary:
    29,138 original records
    29,137 records loaded
         1 record NOT loaded (zNoLoad_Station)

---------

2. Statistic table - the most recent DB had 1,377,502 records

 a  1,101 records did not transfer

    1,089 records had ZLS in DataSource field
    These were changed to '00' to match domain entry 'None'.
    The records then sucessfully loaded.

    12 records were for Station StaID 01482900.
    That Station does not exist in the Station table.
    A dummy record in Station was created for that StaID.
    This record will anchor the Stat values.
    All records then loaded into Statistic.
      1 of these 12 records had 'NAD83' for the Statistic value.

    Statistic table summary:
    1,377,502 original records
    1,376,401 records loaded without initial errors
        1,089 records WERE loaded after replacing a Null DataSource with '00' ("none")
           12 records NOT loaded due to invalid Station (zNoLoad_Statistic)

---------

3. Components table - the most recent DB had 7,844 records.

 a  This table has no Primary Key assigned.  
    No valid natural key exists due to relational issues described below.
    DepVarID + ParmID was not unique; other columns vary more in value.
    * A PK should be determined for this table to enforce record uniqueness.
      * Without a declared key the wrong values could be used.
    * The absence of a PK has led to some referential integrity issues.
 
 b  1 record contained only the value 2136 for DepVarID.
    Although this is a valid ID, the other fields were all empty.
    The record was not loaded.
    
 c  205 records contain the value -1 for DepVarID, and it is not valid.
    No such value exists for DepVarID in the DepVars table.			
    These records were not loaded.
    
 d  357 records contain invalid ParmID values.
    These IDs do not exist in the Parameters table.
    A total of 24 different invalid ParmID values are used,
      including the single Null from (b) above.
    Many of these were already included in the DepVarsID -1 item (c) above.
    These records were not loaded.

 e  After loading the valid records, the assumed PK was checked.
    For the loaded records, DepVarID + ParmID is unique.
    A formal PK should be created and the formal relationships
     to parent tables enforced (from DepVars and Parameters).
    
    Component table summary:
    7,844 original records
    7,487 records loaded without issues
      357 records NOT loaded (zNoLoad_Components)
    The PK and relationships should be formalized with the loaded data.        

---------    
    
4. Covariance table - the most recent DB had 8,343 records.

 a  This table had no Primary Key.  
    The presumed natural key did not exist due to relational issues described below.
    DepVarID + Row + Col was not unique.
    * The cause was a referential integrity issue.
    
 b  431 records have the invalid DepVarID value of -1.
    These records are invalid and some are duplicated 
        with regard to the DepVarID + Row + Col presumed PK.
    72 more records were invalid for 6 additional DepVarID's:
       2476, 10611, 10612, 10613, 10614, 10615
    That is a total of 503 records with invalid DepVarID's.
    These records were not loaded (zNoLoad_Covariance).

 c  After loading the 7,912 valid records, the assumed PK was checked.
    For the loaded records, DepVarID + Row + Col is now unique.
    A formal PK should be created and the formal relationships
     to a parent table enforced (from DepVars).
     
    Covariance table summary:
    8,343 original records   
    7,912 records loaded without issues
      431 records NOT loaded (zNoLoad_Covariance)
    The PK and relationships should be formalized with the loaded data.        
    
---------
    
5. ROIUserParms table - the most recent DB had 32 records.

 a  This table had no Primary Key.  
    There are few records and the table function is unknown.
    
 b  9 records have invalid ParmID values of either 7, 8, or 80
    These were not loaded (zNoLoad_ROIUserParms)

    ROIUserParms table summary:
    32 records loaded without issues
     9 records NOT loaded (zNoLoad_Covariance)
    The PK and relationships should be formalized with the loaded data.        

 
    
---------------------------------

Other Issues

---------

Field Descriptions

  The table DepVars has a Field named "AccessFlag".
  This Field does not have a definition.

--------- 

Extra Fields

  Five (5) Fields were added to the Station table to assist NHD handling and without definitions.
  Recent emails suggest the NHD coordinate and related fields will be handled in a separate table.
  If so, these 5 fields should be removed from the Station table:
 
    NHDReach
    NHDPercent
    NHDScale
    CoordinateScale
    CoordinateSource



---------

StationDistrictState table
  This table has invalid StateCode values that are single digits (no leading zero).
  DistrictCode values can also be single digit (no leading zero).
  No enforceable relationship can be made with States.


---------

StationState

The following StaID values are not valid (not in Station table).
Each has one StationState record:
011058758
01492950
01591000b
02081000
02114010
02129530
02141130
03480540
6013500
6015500
6019500
6037500
9223000


---------

Recommended Change

  StreamStatsDB contains 3 'system' tables:
    DBChangeLog
    DetailedLog
    TransactionLog

  To distinguish these from data tables, consider a common prefix to group these tables ('z', '0', 'sys').




--End