"""
Test script to inventory AGOL/Portal items to workbooks.
Workbooks are broken out into applications, webmaps, and 
standalone servies (services not consumed by webmaps or apps)

"""

from arcgis.gis import GIS
import json
import pandas as pd
import datetime
import os


def documentApps(all_content, item_types, excel, sheet):

    # document applications
    values = []
    # populate list- use to record standalone webmaps and services 
    webmaps = []
    feature_map_services = []

    for item in all_content:
        if item_types.get(item["type"]) == "applications":
            create_date = (item.created)/1000
            create_date_f = datetime.datetime.fromtimestamp(create_date).strftime("%m/%d/%Y, %H:%M:%S")
            update_date = item.modified/1000
            update_date_f = datetime.datetime.fromtimestamp(update_date).strftime("%m/%d/%Y, %H:%M:%S")

            map_itemid = None

            if item["type"] ==  'Dashboard':
                app_type = item["type"]
                item_data = item.get_data()
                if "widgets" in item_data.keys():
                    for w in item_data["widgets"]:
                        if w["type"] == "mapWidget":
                            map_itemid = w["itemId"]
                            if map_itemid not in webmaps:
                                webmaps.append(map_itemid)
            elif item["type"] == "Web Mapping Application":
                app_type = item["type"]
                item_data = item.get_data()
                if "values" in item_data.keys():
                    if "webmap" in item_data["values"]:
                        map_itemid = item_data["values"]["webmap"]
                        if map_itemid not in webmaps:
                                webmaps.append(map_itemid)
                elif "map" in item_data.keys():
                    if "itemId" in item_data["map"]:
                        map_itemid = item_data["map"]["itemId"]
                        if map_itemid not in webmaps:
                                webmaps.append(map_itemid)
            elif item["type"] == "Web Experience":
                app_type = item["type"]
                item_data = item.get_data()
                if "dataSources" in item_data.keys():
                    for d in item_data["dataSources"]:
                        if "type" in item_data["dataSources"][d]:
                            if item_data["dataSources"][d]["type"] == "WEB_MAP":
                                map_itemid = item_data["dataSources"][d]["type"]["itemId"]
                                if map_itemid not in webmaps:
                                    webmaps.append(map_itemid)
            elif item["type"] == "StoryMap":
                app_type = item["type"]
                # TODO: determine how webmap is coded in story maps
            elif item["type"] == "Workforce Project":
                app_type = item["type"]
                # TODO: determine how webmap (if applicable) is coded in WF 
            else:
                app_type = item["type"]
                # capture all other app types that do not carry a webmap as NA
                # ex. Survey123 forms 
            
            if map_itemid != None:
                # it was an app with a detectable webmap
                map_item = gis.content.get(map_itemid)
                if map_item != None:
                
                    map_data = map_item.get_data()
                    layers = map_data.get("operationalLayers")
                    if layers:
                        for lyr in layers:
                            item_id = lyr.get("itemId")
                            if item_id not in feature_map_services:
                                feature_map_services.append(item_id)
                            name = lyr.get("title")
                            url = lyr.get("url")
                            if type(url) == str:
                                # map tours cause problems..
                                lyr_index = url.split(r"/")[-1]
                            else:
                                lyr_index = None
                            # app type, app name, app id, webmap name, webmap id, 
                            # layer name, layer itemid, layer url, layer index 
                            values.append([app_type, item.get("title"), item.get("id"), item.get("owner"), item.get("numViews"), create_date_f, update_date_f,
                                    map_item.get("title"), map_itemid,
                                    name, item_id, url, lyr_index])
                    else:
                        # could not access layers
                        values.append([app_type, item.get("title"), item.get("id"),  item.get("owner"), item.get("numViews"), create_date_f, update_date_f,
                                    map_item.get("title"), map_itemid,
                                    "LAYERS NOT FOUND"])
                else:
                    values.append([app_type, item.get("title"), item.get("id"),  item.get("owner"), item.get("numViews"), create_date_f, update_date_f,
                                "MAP NOT ACCESSIBLE", map_itemid])
            else:
                values.append([app_type, item.get("title"), item.get("id"), item.get("owner"), item.get("numViews"), create_date_f, update_date_f,
                            "NO MAP"])

    df = pd.DataFrame(values, 
                        columns=["App Type", "App Name", "App ID", "App Owner", "View Count", "Created Date", "Updated Date",
                                "WebMap Name", "WebMap Item ID",
                                "Layer Name", "Layer Item ID", "Feature Service URL", "Layer Index"])

    df.to_excel(excel, sheet_name=sheet)

    writer.close()

    return webmaps, feature_map_services


def documentWebmaps(all_content, webmaps, feature_map_services, item_types, excel, sheet):

    values = []
    for item in all_content:
        if item_types.get(item["type"]) == 'maps':
            map_id = item.get("id")
            if map_id not in webmaps:
                item_type = item.get("type")
                create_date = (item.created)/1000
                create_date_f = datetime.datetime.fromtimestamp(create_date).strftime("%m/%d/%Y, %H:%M:%S")
                update_date = item.modified/1000
                update_date_f = datetime.datetime.fromtimestamp(update_date).strftime("%m/%d/%Y, %H:%M:%S")
                map_name = item.get("title")
                
                owner = item.get("owner")
                view_count = item.get("numViews")
                map_item = gis.content.get(map_id)
                if map_item != None:
                
                    map_data = map_item.get_data()
                    layers = map_data.get("operationalLayers")
                    if layers:
                        for lyr in layers:
                            item_id = lyr.get("itemId")
                            if item_id not in feature_map_services:
                                feature_map_services.append(item_id)
                            name = lyr.get("title")
                            url = lyr.get("url")
                            if type(url) == str:
                                # map tours cause problems..
                                lyr_index = url.split(r"/")[-1]
                            else:
                                lyr_index = None
                            values.append([item_type, map_name, map_id, owner, view_count, create_date_f, update_date_f,
                                        name, item_id, url, lyr_index])

    df = pd.DataFrame(values, 
                        columns=["Map Type", "Map Name", "Map ID", "Map Owner", "View Count", "Created Date", "Updated Date",
                                "Layer Name", "Layer Item ID", "Feature Service URL", "Layer Index"])
    df.to_excel(excel, sheet_name=sheet)

    excel.close()

    return feature_map_services


def documentServices(all_content, feature_map_services, item_types, excel, sheet):

    values = []
    for item in all_content:
        if item_types.get(item["type"]) == 'layers':
            item_id = item.get("id")
            if item_id not in feature_map_services:
                item_type = item.get("type")
                name = item.get("title")
                url = item.get("url")
                owner = item.get("owner")
                create_date = (item.created)/1000
                create_date_f = datetime.datetime.fromtimestamp(create_date).strftime("%m/%d/%Y, %H:%M:%S")
                update_date = item.modified/1000
                update_date_f = datetime.datetime.fromtimestamp(update_date).strftime("%m/%d/%Y, %H:%M:%S")
                map_name = item.get("title")
                view_count = item.get("numViews")
                layers = item.get("layers")
                item_data = item.get_data()
                if layers:
                    for lyr in layers:
                        properties = lyr.properties
                        lyr_name = properties.get("name")
                        lyr_id = properties.get("id")
                        values.append([item_type, name, item_id, owner, url, view_count, create_date_f, update_date_f,
                            lyr_name, lyr_id])
                elif item_data.get("layers"):
                    for lyr in item_data.get("layers"):
                        lyr_name = lyr.get("name")
                        lyr_id = lyr.get("id")
                        values.append([item_type, name, item_id, owner, url, view_count, create_date_f, update_date_f,
                            lyr_name, lyr_id])
                else:
                    lyr_id = None
                    lyr_name = None
                    values.append([item_type, name, item_id, owner, url, view_count, create_date_f, update_date_f,
                            lyr_name, lyr_id])
                

    df = pd.DataFrame(values, 
                        columns=["Service Type", "Service Name", "Service ID", "Service Owner", "Service URL", "View Count", "Created Date", "Updated Date",
                                "Layer Name", "Layer Index"])
    df.to_excel(excel, sheet_name=sheet)

    excel.close()


if __name__ == "__main__":

    # AGOL/Portal
    # replace with Portal URL as needed
    ##### POPULATE WITH YOUR VALUES
    gis = GIS("https://arcgis.com", "<username>", "<password>")
    print(gis.properties.user.username)

    # get supplemental 
    home_fldr = os.path.dirname(__file__)
    supp_fldr = os.path.join(home_fldr, "supp")
    
    all_content_json = os.path.join(supp_fldr, "AGO_items_by_group.json")
    item_txt = open(all_content_json).read()
    item_types = json.loads(item_txt)

    all_content = []
    all_users = gis.users.search(max_users=3000)
    for user in all_users:
        name = user.username
        items = gis.content.search(query=f"owner: {name}", max_items=3000)
        for item in items:
            all_content.append(item)

    # TODO: just add sheet, no need for separate workbook per inventory 
    # populate with folder path
    ##### POPULATE WITH YOUR VALUES
    workbook_paths = r""

    # applications
    print("inventorying applications...")
    excel_path = os.path.join(workbook_paths, "inventory_apps.xlsx")
    curr_sheet_name = "AGOL Inventory"
    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")

    webmaps_list, feature_services_list = documentApps(all_content, item_types, writer, curr_sheet_name)

    # webmaps
    print("inventorying webmaps...")
    excel_path = os.path.join(workbook_paths, "webmap_inventory.xlsx")
    curr_sheet_name = "AGOL Inventory"
    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")

    feature_services_list_u = documentWebmaps(all_content, webmaps_list, feature_services_list, item_types, writer, curr_sheet_name)

    # feature services 
    print("inventorying services...")
    excel_path = os.path.join(workbook_paths, "service_inventory.xlsx")
    curr_sheet_name = "AGOL Inventory"
    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")

    documentServices(all_content, feature_services_list_u, item_types, writer, curr_sheet_name)

