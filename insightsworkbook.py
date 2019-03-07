""" Class for interacting with ArcGIS Insights """

import json
import random
import sys
from urllib.error import HTTPError

from arcgis.gis import GIS

__version__ = '0.1'


class InsightsWorkbook(object):
    """ An object representing an ArcGIS Insights workbook

    Users can use this class to create new Insights workbooks, connect to
    existing workbooks, and modify them. Currently, it serves as a proof of
    concept to demonstrate adding data, creating a map, and creating a chart,
    but new functionality can easily be added in the future.


    ================    ========================================================
    **Argument**        **Description**
    ----------------    --------------------------------------------------------
    gis                 GIS connection object from arcgis.gis (ArcGIS API for
                        Python). Connection should be set up before working with
                        this class.
    ----------------    --------------------------------------------------------
    title               Optional string. Title for the Insights workbook
    ----------------    --------------------------------------------------------
    workbook_id         Optional string. Random ID used for Hosted storage.
    ----------------    --------------------------------------------------------
    workspace_id        Optional string. Workbook Item ID in ArcGIS - generated
                        by ArcGIS upon creation of a workbook.
    ----------------    --------------------------------------------------------
    workspace_url       Optional string. URL where the Insights workbook
                        internal hosted workspace resides. Structure differs
                        between ArcGIS Online and ArcGIS Portal.
    ----------------    --------------------------------------------------------
    props               Optional dict. Dictionary of workbook properties that
                        represents the full JSON data object that is stored in
                        ArcGIS.
    ================    ========================================================




    .. code-block:: python

        # Usage Example 1: Anonymous Login to ArcGIS Online
    """

    def __init__(self, gis, title=None, workbook_id=None, workspace_id=None,
                 workspace_url=None, props=None):
        """
        Constructs the Workbook given the aforementioned parameters. Normally,
        the Workbook object will be created with either the new() or open()
        class methods below.
        """
        self._gis = gis
        self._title = title
        self._agol = False
        # Check whether it's AGOL or Portal
        if gis._url.find('arcgis.com') >= 0:
            self._agol = True
        self._workbookID = workbook_id
        self._workspaceID = workspace_id
        self._workspaceURL = workspace_url
        self.props = props

    @classmethod
    def new(cls, gis, title):
        """
        Creates a new Insights Workbook in ArcGIS using the provided title.

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        gis                    Required arcgis.gis.GIS. Connection should be set
                               up before working with this class.
        ------------------     -------------------------------------------------
        title                  Optional string. Title to be assigned to the new
                               Workbook.
        ==================     =================================================

        :return:
           New InsightsWorkbook object with the provided title.
        """
        # Random 8-digit hex number for ID
        workbook_id = '%08x' % random.randrange(16**8)

        # Hosted service URL differs between Portal and AGOL
        insights_service_url = gis._url.lower() + \
            '/arcgis/rest/services/Hosted/'
        if gis._url.find('arcgis.com') >= 0:
            insights_service_url = 'https://insightsservices.arcgis.com/' +\
                                    gis._portal._properties['id'] + \
                                    '/arcgis/rest/services/'
        workspace_url = insights_service_url + workbook_id + "/WorkspaceServer"

        # Path for REST API to create new Workspace Service (ArcGIS Insights
        # internal data storage)
        path = gis._portal.url + '/sharing/rest/content/users/' + \
            gis.users.me.username + '/createService'
        post_data = {'f': 'json',
                     'createParameters': '{"name": "' + workbook_id + '"}',
                     'targetType': 'workspaceService'}
        # Catch any errors from POSTing
        try:
            # The first call creates the workspace, but the Workbook is not
            # functional until after the second POST below
            resp = gis._portal.con.post(path, post_data)
            # Depending on whether it's on Portal or AGOL, it will use a
            # different key for ID.
            if gis._url.find('arcgis.com') < 0:
                workspace_id = resp['itemId']
            else:
                workspace_id = resp['serviceItemId']

            item_props = {
                "f": "json",
                "id": workspace_id,
                "type": "Insights Workbook",
                "title": title}
            # For some reason this is not needed for Portal, but it is for AGOL
            if gis._url.find('arcgis.com') >= 0:
                item_props["url"] = workspace_url

            # Full set of default JSON data properties for an ArcGIS Insights
            # workbook item.
            props = {
                "format": 9,
                "title": title,
                "pages": [{
                    "title": "Page 1",
                    "model": {
                        "items": []
                    },
                    "cards": [],
                    "layout": [],
                    "contents": []
                }],
                "activePage": 0,
                "workspace": {
                    "datasets": {}
                },
                "_ssl": True,
                "created": 0,
                "modified": 0,
                "guid": None,
                "type": "Insights Workbook",
                "typeKeywords": ["Application", "ArcGIS",
                                 "Insights Workbook", "Hosted Service"],
                "description": None,
                "tags": [],
                "snippet": None,
                "thumbnail": None,
                "documentation": None,
                "extent": [],
                "categories": [],
                "spatialReference": None,
                "accessInformation": None,
                "licenseInfo": None,
                "culture": "english (united states)",
                "properties": None,
                "proxyFilter": None,
                "access": "private",
                "size": 0,
                "appCategories": [],
                "industries": [],
                "languages": [],
                "largeThumbnail": None,
                "banner": None,
                "screenshots": [],
                "listed": False,
                "ownerFolder": None,
                "protected": False,
                "commentsEnabled": True,
                "numComments": 0,
                "numRatings": 0,
                "avgRating": 0,
                "numViews": 3,
                "itemControl": "admin",
                "scoreCompleteness": 0,
                "groupDesignations": None
            }
            # After the first call sets up the workspace, this second call sets
            # up the actual Workbook with all the data props, title, etc.
            gis._portal.update_item(workspace_id, item_props,
                                           json.dumps(props))
            # Now that it's created, store the relevant properties in this class
            # for use in other functions (e.g. add data, create map, etc.)
            return cls(gis, title, workbook_id,
                       workspace_id, workspace_url, props)
        except:
            print("Error creating workbook: " + sys.exc_info()[0])
            raise Exception()

    @classmethod
    def open(cls, existing_workbook):
        """
        Creates a new Insights Workbook in ArcGIS using the provided title.

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        gis                    Required arcgis.gis.Item. Must pass a previously-
                               created ArcGIS Insights Workbook. The ArcGIS API
                               for Python Item object should be of the
                               "Insights Workbook" type.
        ==================     =================================================

        :return:
           InsightsWorkbook object that points to this existing Workbook
        """
        gis = existing_workbook._gis
        title = existing_workbook.title
        workbook_id = existing_workbook.name
        # Hosted service URL differs between Portal and AGOL
        insights_service_url = gis._url.lower() + \
            '/arcgis/rest/services/Hosted/'
        if gis._url.find('arcgis.com') >= 0:
            insights_service_url = 'https://insightsservices.arcgis.com/' +\
                                    gis._portal._properties['id'] + \
                                    '/arcgis/rest/services/'
        workspace_url = insights_service_url + workbook_id + "/WorkspaceServer"
        workspace_id = existing_workbook.id
        # Path to get the JSON data for this Workbook item
        path = gis._url.lower() + '/sharing/rest/content/items/' + \
            workspace_id + '/data'
        try:
            resp = gis._portal.con.get(path, {'f': 'json'})
            props = resp
            return cls(gis, title, workbook_id,
                       workspace_id, workspace_url, props)
        except:
            print("Error retrieving workbook data: " + sys.exc_info()[0])
            raise Exception()

    def add_feature_layer(self, lyr, sublayer=0):
        """
        Adds a feature layer as a dataset to this Workbook

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        lyr                    Required arcgis.gis.Item. Must pass a feature
                               layer from the ArcGIS API for Python.
        ------------------     -------------------------------------------------
        sublayer               Optional int. Index of the sublayer to add.
        ==================     =================================================
        :return:
           String name of new internal Insights Workbook dataset
        """
        # Random 7-digit hex suffix for dataset name
        dataset_name = lyr.title + '_%07x' % random.randrange(16**7)
        post_data = {'f': 'json',
                     'tools': '[{"name":"add-data","params":{"data":{' +
                              '"type":"feature-layer",' +
                              '"url":"' + lyr.url + '/' + str(sublayer) +
                              '"}},"outDataset":"' + dataset_name + '"}]',
                     'outDatasets': '["' + dataset_name + '"]'}
        try:
            # Execute the add-data operation within ArcGIS Insights. Note: data
            # is not automatically saved in the Workbook. Must manually call
            # save in order to make it permanent.
            resp = self._gis._portal.con.post(self._workspaceURL + '/execute',
                                              post_data)

            # In addition to calling execute to create the dataset, it must also
            # be placed in the Workbook Item properties in several places:
            self.props['pages'][0]['model']['items'].append({
                'operation': 'add-data',
                'params': {
                    'data': {
                        'type': 'feature-layer',
                        'url': lyr.url + '/' + str(sublayer)
                    }
                },
                'outDataset': dataset_name
            })
            self.props['pages'][0]['contents'].append({
                'dataset': dataset_name
            })
            my_extent = json.loads(lyr.layers[sublayer].properties.extent.__repr__())
            self.props['workspace']['datasets'][dataset_name] = {
                'data': resp[dataset_name],
                'owner': lyr.id,
                'fields': {
                    'shape': {
                        'alias': 'Location'
                    }
                },
                'extent': my_extent,
                'origin': True
            }
            return dataset_name
        except:
            print("Error adding feature layer: " + sys.exc_info()[0])
            raise Exception()

    def update_dataset(self, lyr, sublayer=0):
        """
        Updates all references to the provided feature layer within this
        Workbook

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        lyr                    Required arcgis.gis.Item. Must pass a feature
                               layer from the ArcGIS API for Python.
        ------------------     -------------------------------------------------
        sublayer               Optional int. Index of the sublayer to add.
        ==================     =================================================
        :return:
           String name of internal Insights Workbook dataset
        """
        # Get the current Workbook model items
        model_items = self.props['pages'][0]['model']['items']
        # Find the specific model item related to adding this dataset
        add_op = [x for x in model_items if ('outDataset' in x and
            'params' in x and
            'data' in x['params'] and
            'url' in x['params']['data'] and
            x['params']['data']['url'].find(lyr.url) >= 0)]
        # Check to see if this layer actually exists in this workbook
        if len(add_op) > 0:
            # Get internal name of this dataset
            dataset_name = add_op[0]['outDataset']
            post_data = {'f': 'json',
                         'tools': '[{"name":"add-data","params":{"data":{' +
                                  '"type":"feature-layer",' +
                                  '"url":"' + add_op[0]['params']['data']['url'] +
                                  '"}},"outDataset":"' + dataset_name + '"}]',
                         'outDatasets': '["' + dataset_name + '"]'}
            try:
                # Execute the add-data operation within ArcGIS Insights. This
                # runs identically to the add data operation - it just
                # generates a new data ID for us to work with.
                resp = self._gis._portal.con.post(self._workspaceURL +
                                                  '/execute', post_data)
                old_data = self.props['workspace']['datasets'][dataset_name]['data']
                # Update existing workspace entry in JSON with new data ID
                my_extent = json.loads(lyr.layers[sublayer].properties.extent.__repr__())
                self.props['workspace']['datasets'][dataset_name] = {
                    'data': resp[dataset_name],
                    'owner': lyr.id,
                    'fields': {
                        'shape': {
                            'alias': 'Location'
                        }
                    },
                    'extent': my_extent,
                    'origin': True
                }
                # Find all references to old data ID and replace with new
                for ds in self.props['workspace']['datasets'].keys():
                    try:
                        ds_tools = self.props['workspace']['datasets'][ds]['data']['tools']
                        for i in range(len(ds_tools)):
                            if self.props['workspace']['datasets'][ds]['data']['tools'][i]['params']['dataset'] == old_data:
                                self.props['workspace']['datasets'][ds]['data']['tools'][i]['params']['dataset'] = resp[dataset_name]
                    except (KeyError, TypeError):
                        # If particular dataset doesn't have this key, move to
                        # next.
                        continue
                # Send back the dataset name
                return dataset_name
            except HTTPError:
                print("Error updating feature layer: " + sys.exc_info()[0])
        else:
            print("Layer does not exist within this Workbook")
            raise Exception()

    def add_map(self, dataset):
        """
        Adds a map card for the specified feature layer dataset

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        dataset                Required string. Internal dataset name for the
                               feature layer to add to the map. Can get this
                               value when feature layer is added to the
                               Workbook. Alternatively, can look it up in the
                               props.
        ==================     =================================================
        """
        # Grab full dataset info from the props
        layer_dataset = self.props['workspace']['datasets'][dataset]
        if layer_dataset:
            # Get current number of cards
            card_ct = len(self.props['pages'][0]['cards'])
            # Get extent of this layer
            my_extent = self.props['workspace']['datasets'][dataset]['extent']
            # Add map, with layer and extent
            self.props['pages'][0]['cards'].append({
                'type': 'map',
                'title': 'Card ' + str(card_ct+1),
                'content': {
                    'layers': [{'datasetId': dataset}],
                    'extent': my_extent
                }
            })
            # Just place new map to the right of the last card
            layout_x = card_ct * 20 + card_ct
            self.props['pages'][0]['layout'].append({
                "x": layout_x,
                "y": 0,
                "w": 20,
                "h": 20
            })
        else:
            print("Invalid dataset name")
            raise Exception()

    def aggregate(self, in_dataset, groupby_field, groupby_field_type,
                  stat_type, stat_field, stat_field_type, out_name=None):
        """
        Aggregates data based on the group-by layer using the specified
        statistic over the specified field.

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        in_dataset             Required str. Internal name of dataset to use for
                               aggregating.
        ------------------     -------------------------------------------------
        groupby_field          Required str. Name of the field to group by.
        ------------------     -------------------------------------------------
        groupby_field_type     Required str. Type of field. Acceptable values
                               include esriFieldTypeString, esriFieldTypeDouble,
                               esriFieldTypeInteger, etc.
        ------------------     -------------------------------------------------
        stat_type              Required str. Type of aggregation statistic.
                               Acceptable values include avg, sum, count, min
                               and max.
        ------------------     -------------------------------------------------
        stat_field             Required str. Name of the field to use in
                               calculation.
        ------------------     -------------------------------------------------
        stat_field_type        Required str. Type of field for output.
                               Acceptable values include esriFieldTypeString,
                               esriFieldTypeDouble, esriFieldTypeInteger, etc.
        ------------------     -------------------------------------------------
        out_name               Optional str. Name of dataset to show to user.
        ==================     =================================================
        :return:
           String name of new internal Insights Workbook dataset for this
           aggregation
        """
        # Get base name of dataset so we can generate a new suffix for new
        # aggregate dataset
        in_dataset_base = in_dataset[:in_dataset.find('_')]
        # Generate new dataset name with 7 random hex digits as suffix
        out_dataset = in_dataset_base + '_%07x' % random.randrange(16**7)
        # New field for storing aggregation
        out_field = stat_field.lower() + '_' + stat_type
        if stat_type == 'count':
            out_type = 'esriFieldTypeInteger'
        elif stat_type == 'avg':
            out_type = 'esriFieldTypeDouble'
        else:
            out_type = stat_field_type
        # Perform aggregate operation
        self.props['pages'][0]['model']['items'].append({
            'operation': 'aggregate',
            'params': {
                'dataset': in_dataset,
                'groupBy': [groupby_field],
                'statistics': [{
                    'type': stat_type,
                    'field': stat_field,
                    'outField': out_field
                }],
                'totals': False
            },
            'outDataset': out_dataset
        })
        in_data_id = self.props['workspace']['datasets'][in_dataset]['data']
        # If no name specified for this dataset just use internal ID
        if not out_name:
            out_name = out_dataset
        # Add new dataset to workspace
        self.props['workspace']['datasets'][out_dataset] = {
            'data': {
                'metadata': {
                    'fields': [
                        {
                            'name': groupby_field,
                            'alias': groupby_field,
                            'type': groupby_field_type,
                            'entity': 'e0'
                        },
                        {
                            'name': out_field,
                            'alias': out_field,
                            'type': out_type
                        }
                    ],
                    'entities': [{
                        'fields': [
                            groupby_field,
                            out_field
                        ],
                        'keyFields': [groupby_field]
                    }]
                },
                'tools': [{
                    'name': 'aggregate',
                    'params': {
                        'dataset': in_data_id,
                        'groupBy': [groupby_field],
                        'statistics': [{
                            'type': stat_type,
                            'field': stat_field,
                            'outField': out_field
                        }],
                        'materialize': False
                    },
                    'outDataset': out_dataset
                }],
                'outDataset': out_dataset
            },
            'name': out_name,
            'fields': {
                out_field: {
                    'alias': stat_type.title() + " of " + in_dataset_base
                },
                groupby_field: {}
            }
        }
        return out_dataset

    def add_chart(self, chart_type, in_dataset, groupby_field,
                  groupby_field_type, stat_type, stat_field, stat_field_type):
        """
        Aggregates data based on the group-by layer using the specified
        statistic over the specified field.

        ==================     =================================================
        **Argument**           **Description**
        ------------------     -------------------------------------------------
        chart_type             Required str. Type of chart to create. Acceptable
                               values are bar and column.
        ------------------     -------------------------------------------------
        in_dataset             Required str. Internal name of dataset to use for
                               aggregating.
        ------------------     -------------------------------------------------
        groupby_field          Required str. Name of the field to group by.
        ------------------     -------------------------------------------------
        groupby_field_type     Required str. Type of field. Acceptable values
                               include esriFieldTypeString, esriFieldTypeDouble,
                               esriFieldTypeInteger, etc.
        ------------------     -------------------------------------------------
        stat_type              Required str. Type of aggregation statistic.
                               Acceptable values include avg, sum, count, min
                               and max.
        ------------------     -------------------------------------------------
        stat_field             Required str. Name of the field to use in
                               calculation.
        ------------------     -------------------------------------------------
        stat_field_type        Required str. Type of field for output.
                               Acceptable values include esriFieldTypeString,
                               esriFieldTypeDouble, esriFieldTypeInteger, etc.
        ==================     =================================================
        """
        # Get count of cards
        card_ct = len(self.props['pages'][0]['cards'])
        # Set chart title
        chart_title = chart_type.title() + " Chart " + str(card_ct+1)
        # Create aggregation dataset for this chart
        out_dataset = self.aggregate(in_dataset, groupby_field,
                                     groupby_field_type, stat_type, stat_field,
                                     stat_field_type, chart_title)
        # Create the chart card
        self.props['pages'][0]['cards'].append({
            'title': 'Card ' + str(card_ct+1),
            'type': 'chart',
            'content': {
                'layers': [{
                    'datasetId': out_dataset,
                    'chartLines': {
                        'mean': True
                    },
                    'colors': {}
                }],
                'type': chart_type
            }
        })
        # Place it to the right of any existing cards
        layout_x = card_ct * 20 + card_ct
        self.props['pages'][0]['layout'].append({
            "x": layout_x,
            "y": 0,
            "w": 20,
            "h": 20
        })

    def save(self):
        """
        Saves the Insights Workbook to ArcGIS with all the current properties.
        """
        # A few properties have to be manually set at save (doesn't work to
        # just set them on initial Workbook creation).
        self.props["id"] = self._workspaceID
        self.props["owner"] = self._gis.users.me.username
        self.props["name"] = self._workbookID
        self.props["url"] = self._workspaceURL
        post_data = {
            'f': 'json',
            'title': self._title,
            'text': json.dumps(self.props)}
        # Basically just a standard ArcGIS item update with the updated JSON
        # properties
        update_url = self._gis._url.lower() + '/sharing/rest/content/users/' + \
            self._gis.users.me.username + '/items/' + \
            self._workspaceID + '/update'
        try:
            self._gis._portal.con.post(update_url, post_data)
        except:
            print("Error saving workspace: " + sys.exc_info()[0])
            raise Exception()
