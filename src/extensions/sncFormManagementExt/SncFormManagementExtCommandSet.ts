import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  // type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { override } from "@microsoft/decorators";
import { Constants } from "./Models/Constants";

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

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SncFormManagementExtCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
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
    let LibraryName = this.context.pageContext.list?.title;

    //* Check if current library name is in Constants.BUSINESS_TRAVEL_LIBRARY_NAME
    if (Constants.BUSINESS_TRAVEL_LIBRARY_NAME.indexOf(LibraryName || "") !== -1) {
      BusinessTravelCompareOneCommand.visible = true;
      if (BusinessTravelSecondCommand) {
        BusinessTravelSecondCommand.visible = event.selectedRows?.length === 1;
      }
      require("./ExtStyle/ExtStyle.css");
    } else {
      BusinessTravelCompareOneCommand.visible = false;
      BusinessTravelSecondCommand.visible = false;
    }

    //* Check if current library name is in Constants.MEETING_ROOM_LIBRARY_NAME
    if (Constants.MEETING_ROOM_LIBRARY_NAME.indexOf(LibraryName || "") !== -1) {
      MeetingRoomCompareOneCommand.visible = true;
      if (MeetingRoomCompareSecondCommand) {
        MeetingRoomCompareSecondCommand.visible = event.selectedRows?.length === 1;
      }
      require("./ExtStyle/ExtStyle.css");
    } else {
      MeetingRoomCompareOneCommand.visible = false;
      MeetingRoomCompareSecondCommand.visible = false;
    }

  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const webUri = this.context.pageContext.web.absoluteUrl;

    switch (event.itemId) {
      //* Create and Update BusinessTravel Request
      case "New_BusinessTravel_Field":
        window.location.href = webUri + Constants.newBusinessTravelRequest;
        break;
      case "Edit_BusinessTravel_Field":
        var ItemID = event.selectedRows[0].getValueByName("ID").toString();
        window.location.href =
          webUri + Constants.editBusinessTravelRequest + "?FormID=" + ItemID;
        break;

      //* Create and Update MeetingRoom Request
      case "New_MeetingRoom_Field":
        window.location.href = webUri + Constants.newMeetingRoomRequest;
        break;
      case "Edit_MeetingRoom_Field":
        var ItemID = event.selectedRows[0].getValueByName("ID").toString();
        window.location.href =
          webUri + Constants.editMeetingRoomRequest + "?FormID=" + ItemID;
        break;

      default:
        throw new Error('Unknown command');
    }
  }
}
