import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  // type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { override } from "@microsoft/decorators";
import { Constants } from "./Models/Constants";
import './ExtStyle/GlobalStyle.css'

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISncFormManagementExtCommandSetProperties {
  // This is an example; replace with your own properties
  New_BusinessTravel_Field: string;
  Edit_BusinessTravel_Field: string;

  New_MeetingRoom_Field: string;
  Edit_MeetingRoom_Field: string;
}

const LOG_SOURCE: string = 'SncFormManagementExtCommandSet';

export default class SncFormManagementExtCommandSet extends BaseListViewCommandSet<ISncFormManagementExtCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SncFormManagementExtCommandSet');
    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewUpdated);

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(): void {

    //* Create and Update Business Travel Request
    const BusinessTravelCompareOneCommand: Command = this.tryGetCommand(
      "New_BusinessTravel_Field"
    );
    const BusinessTravelSecondCommand: Command = this.tryGetCommand(
      "Edit_BusinessTravel_Field"
    );

    //* Create and Update Meeting Room Request
    const MeetingRoomCompareOneCommand: Command = this.tryGetCommand(
      "New_MeetingRoom_Field"
    );
    const MeetingRoomCompareSecondCommand: Command = this.tryGetCommand(
      "Edit_MeetingRoom_Field"
    );

    //* Get current library name from pageContext
    let LibraryName = this.context.pageContext.list?.title as string;
    console.log("LibraryName:", LibraryName)

    //* Check if current library name is in Constants.BUSINESS_TRAVEL_LIBRARY_NAME
    if (Constants.BUSINESS_TRAVEL_LIBRARY_NAME.indexOf(LibraryName) !== -1) {
      BusinessTravelCompareOneCommand.visible = true
      if (BusinessTravelSecondCommand) {
        BusinessTravelSecondCommand.visible = this.context.listView.selectedRows?.length === 1;
      }
      require("./ExtStyle/ExtStyle.css");
    } else {
      BusinessTravelCompareOneCommand.visible = false;
      BusinessTravelSecondCommand.visible = false;
    }

    //* Check if current library name is in Constants.MEETING_ROOM_LIBRARY_NAME
    if (Constants.MEETING_ROOM_LIBRARY_NAME.indexOf(LibraryName) !== -1) {
      MeetingRoomCompareOneCommand.visible = true;
      if (MeetingRoomCompareSecondCommand) {
        MeetingRoomCompareSecondCommand.visible = this.context.listView.selectedRows?.length === 1;
      }
      require("./ExtStyle/ExtStyle.css");
    } else {
      MeetingRoomCompareOneCommand.visible = false;
      MeetingRoomCompareSecondCommand.visible = false;
    }
    this.raiseOnChange();
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const webUri = this.context.pageContext.web.absoluteUrl;
    console.log("webUri:", webUri)

    switch (event.itemId) {
      //* Create and Update BusinessTravel Request
      case "New_BusinessTravel_Field":
        window.location.href = webUri + Constants.newBusinessTravelRequest;
        break;
      case "Edit_BusinessTravel_Field":
        if (!this.context.listView.selectedRows?.length) return
        var ItemID = this.context.listView.selectedRows[0].getValueByName("ID").toString();
        window.location.href =
          webUri + Constants.editBusinessTravelRequest + "?FormID=" + ItemID;
        break;

      //* Create and Update MeetingRoom Request
      case "New_MeetingRoom_Field":
        window.location.href = webUri + Constants.newMeetingRoomRequest;
        break;
      case "Edit_MeetingRoom_Field":
        if (!this.context.listView.selectedRows?.length) return
        var ItemID = this.context.listView.selectedRows[0].getValueByName("ID").toString();
        window.location.href =
          webUri + Constants.editMeetingRoomRequest + "?FormID=" + ItemID;
        break;

      default:
        throw new Error('Unknown command');
    }
  }
}