import * as React from "react";
import { useState, useCallback, useRef, useEffect } from "react";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { IQuoteRequestProps } from "./IQuoteRequestProps";
import { FontSizes } from "@fluentui/theme";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Separator } from "office-ui-fabric-react/lib/Separator";
import { Text } from "office-ui-fabric-react/lib/Text";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import { Icon, loadTheme, createTheme } from "@fluentui/react";


var alertify: any = require("../../../ExternalRef/js/alertify.min.js");
var html2pdf:any=require("../../../ExternalRef/js/html2pdf.bundle.min.js");
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as $ from "jquery";
import { Label } from "@microsoft/office-ui-fabric-react-bundle";
import { Hidden } from "@material-ui/core";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

//import 'office-ui-fabric-react/dist/css/fabric.css';
//import "../../../ExternalRef/css/style.css";
// import '../../../ExternalRef/css/QuoteRequestStyle.css';
//import  styles from "../components/QuoteRequest.module.scss";

import {PeoplesData} from "./PeoplesData";

let getSelectedUsers: number[] = [];
let NextOrderID: string = "";
let Category: String = "";
const dialogContentProps = {
  type: DialogType.normal,
  title: "Form Submitted Successfully",
};
var FormFilled: boolean = false;
interface formvalues {
  DeviceCount: string;
  OrderNo: string;
  Category: string;
  PointCount: string;
  Drivers: string;
  SpecialConsiderations: string;
  BEName: string;
  BENumber: string;
  ManagerName: string;
  ManagerPhoneNumber: string;
  ManagerEmail: string;
  VendorManagerName: string;
  VendorManagerPhoneNumber: string;
  VendorManagerEmail: string;
  ShippingCountName: string;
  ShippingAddress: string;
  AdditionalInformation: string;
  Description: string;
  UserDetailsId: string | Number[] | any;
}
const formRowStyles = { display: "flex", width: "100%",marginTop:"0.5rem" };
const formColStyles = { width: "23%", padding: "0rem 1rem" };
function RequestNewQuoteAdmin(
  props: IQuoteRequestProps
): React.ReactElement<[]> {
  const wellsFargoTheme = createTheme({
    palette: {
      themePrimary: "#d71e2b",
      themeLighterAlt: "#fdf5f5",
      themeLighter: "#f8d6d9",
      themeLight: "#f3b4b8",
      themeTertiary: "#e77078",
      themeSecondary: "#db3540",
      themeDarkAlt: "#c11b26",
      themeDark: "#a31720",
      themeDarker: "#781118",
      neutralLighterAlt: "#faf9f8",
      neutralLighter: "#f3f2f1",
      neutralLight: "#edebe9",
      neutralQuaternaryAlt: "#e1dfdd",
      neutralQuaternary: "#d0d0d0",
      neutralTertiaryAlt: "#c8c6c4",
      neutralTertiary: "#a19f9d",
      neutralSecondary: "#605e5c",
      neutralPrimaryAlt: "#3b3a39",
      neutralPrimary: "#323130",
      neutralDark: "#201f1e",
      black: "#000000",
      white: "#ffffff",
    },
  });
  loadTheme(wellsFargoTheme);
  let siteURL = props.context.pageContext.web.absoluteUrl;

  const intialvalues: formvalues[] = [
    {
      DeviceCount: "",
      OrderNo: "",
      Category: "",
      PointCount: "",
      Drivers: "",
      SpecialConsiderations: "",
      BEName: "",
      BENumber: "",
      ManagerName: "",
      ManagerPhoneNumber: "",
      ManagerEmail: "",
      VendorManagerName: "",
      VendorManagerPhoneNumber: "",
      VendorManagerEmail: "",
      ShippingCountName: "",
      ShippingAddress: "",
      AdditionalInformation: "",
      Description: "",
      UserDetailsId: "",
    },
  ];

  const intialvalidations: formvalues[] = [
    {
      DeviceCount: "",
      OrderNo: "",
      Category: "",
      PointCount: "",
      Drivers: "",
      SpecialConsiderations: "",
      BEName: "",
      BENumber: "",
      ManagerName: "",
      ManagerPhoneNumber: "",
      ManagerEmail: "",
      VendorManagerName: "",
      VendorManagerPhoneNumber: "",
      VendorManagerEmail: "",
      ShippingCountName: "",
      ShippingAddress: "",
      AdditionalInformation: "",
      Description: "",
      UserDetailsId: "",
    },
  ];

  //const arrVlidations:formvalues[]=intialvalidations;

  //const arrValues:formvalues[]=intialvaluestemp;

  const [Column, setColumn] = useState(true);
  const [Hidedialog, setHidedialog] = useState(true);
  ///const[FormFilled,setFormFilled]=useState(false);

  //const[Validation,setValidation]=useState(arrVlidations);
  //const[Submitvalues,setSubmitvalues]=useState(arrValues);

  const [Validation, setValidation] = useState<formvalues[] | undefined>(
    intialvalidations
  );
  const [Submitvalues, setSubmitvalues] = useState<formvalues[] | undefined>(
    intialvalues
  );
  const [Selectedpeoples,setSelectedpeoples]=useState([]);
  const [Groups, setGroups] = useState("");

  useEffect(() => {
    getLastID();
    setGroups("LynxSpring Members");
    var queryParms = new UrlQueryParameterCollection(window.location.href);
    var myParm = queryParms.getValue("Category");
    if (myParm) {
      if (myParm.toLowerCase() == "admin") Category = "Admin";
      else Category = "Retail";
    } else {
      location.href = siteURL + "/SitePages/WFRequestDashboard.aspx";
    }
    setTimeout(() => {
      setColumn(false);
    }, 2000);
  }, []);

  return (
    <div>
      <div id="adminFormRequest" style={{margin:"1rem 2rem"}}>
        {Column && (
          <Spinner
            label="Loading items..."
            size={SpinnerSize.large}
            style={{
              width: "100vw",
              height: "100vh",
              position: "fixed",
              top: 0,
              left: 0,
              backgroundColor: "#fff",
              zIndex: 10000,
            }}
          />
        )}
        <div className="ms-Grid">
          <div style={formRowStyles}>
            <div>
              <div className={"txtClassArrow"}>
                <Icon
                  iconName="NavigateBack"
                  onClick={DialogBox}
                  styles={{
                    root: {
                      color: wellsFargoTheme.palette.themePrimary,
                      fontSize: "2rem",
                      cursor: "pointer",
                    },
                  }}
                />
              </div>
            </div>
            <div style={{ margin: "auto" }}>
              <Text
                variant={"xLarge"}
                className={"txtClassHead"}
                styles={{
                  root: {
                    color: wellsFargoTheme.palette.themePrimary,
                    fontWeight: "bold",
                    fontSize: "1.5rem",
                    textAlign: "center",
                    width: "100%",
                  },
                }}
              >
                {Category == "Admin"
                  ? "Admin Form Request"
                  : "Retail Form Request"}
              </Text>
            </div>
          </div>
          <div className="html2pdf__page-break"></div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <TextField
                styles={{ root: { width: "100%" } }}
                label="Device count"
                id="txtDeviceCount"
                name={"DeviceCount"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].DeviceCount}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <div className="ms-Grid" dir="ltr">
                <TextField
                  label="Point count"
                  id="txtPointCount"
                  name={"PointCount"}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].PointCount}
                ></TextField>
              </div>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Drivers"
                id="txtDrivers"
                name={"Drivers"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].Drivers}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Site Name"
                id="txtBEName"
                name={"BEName"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].BEName}
              ></TextField>
            </div>
          </div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <TextField
                label="BE Number"
                id="txtBENumber"
                name={"BENumber"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].BENumber}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                multiline
                rows={2}
                resizable={false}
                autoAdjustHeight
                label="Special Considerations"
                id="txtSpecialConsiderations"
                name={"SpecialConsiderations"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].SpecialConsiderations}
              ></TextField>
            </div>
          </div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <Text variant={"xLarge"} className={"txtClassHead"} style={{marginTop:"0.5rem"}}>
                Project Manager Details:
              </Text>
            </div>
          </div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <TextField
                label="Name"
                id="txtManagerName"
                name={"ManagerName"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].ManagerName}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Email - ID"
                id="txtManagerEmailID"
                name={"ManagerEmail"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].ManagerEmail}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Phone Number"
                id="txtManagerPhoneNumber"
                name={"ManagerPhoneNumber"}
                onChange={(e) => handlechange(e)}
                required
                type="number"
                errorMessage={Validation[0].ManagerPhoneNumber}
              ></TextField>
            </div>
          </div>
          <div style={formRowStyles}>
            
          </div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <Text variant={"xLarge"} className={"txtClassHead"}>
                Vendor Manager Details:
              </Text>
            </div>
          </div>

          <div style={formRowStyles}>
            <div style={formColStyles}>
              <TextField
                label="Name"
                id="txtVendorManagerName"
                name={"VendorManagerName"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].VendorManagerName}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Email - ID"
                id="txtVendorManagerEmailID"
                name={"VendorManagerEmail"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].VendorManagerEmail}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                label="Phone Number"
                id="txtVendorManagerPhoneNumber"
                name={"VendorManagerPhoneNumber"}
                onChange={(e) => handlechange(e)}
                required
                type="number"
                errorMessage={Validation[0].VendorManagerPhoneNumber}
              ></TextField>
            </div>
          </div>
          <div style={formRowStyles}>
            
            </div>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <Text variant={"xLarge"} className={"txtClassHead"} style={{marginTop:"0.5rem"}}>
                Shipping Details:
              </Text>
            </div>
          </div>

          <div style={formRowStyles}>
            <div style={formColStyles}>
              <TextField
                label="Shipping Cont Name"
                id="txtShippingCountName"
                name={"ShippingCountName"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].ShippingCountName}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                multiline
                rows={2}
                resizable={false}
                autoAdjustHeight
                label="Shipping Address"
                id="txtShippingAddress"
                name={"ShippingAddress"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].ShippingAddress}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                multiline
                rows={2}
                resizable={false}
                autoAdjustHeight
                label="Additional Information"
                id="txtAdditionalInformation"
                name={"AdditionalInformation"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].AdditionalInformation}
              ></TextField>
            </div>
            <div style={formColStyles}>
              <TextField
                multiline
                rows={2}
                resizable={false}
                autoAdjustHeight
                label="Description"
                id="txtDescription"
                name={"Description"}
                onChange={(e) => handlechange(e)}
                required
                errorMessage={Validation[0].Description}
              ></TextField>
            </div>
          </div>
          <div style={formRowStyles}>
          </div>
          <Separator></Separator>
          <div style={formRowStyles}>
            <div style={formColStyles}>
              <PeoplesData update={UpdateSelectedUsers} spcontext={props.spcontext}/>
              {/*<PeoplePicker
                context={props.context}
                titleText="Add User"
                personSelectionLimit={1}
                groupName={"LynxSpring Members"} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                //required
                //errorMessage={Validation[0]['UserDetailsId']}
                showHiddenInUI={false}
                onChange={(e) => getUserID(e)}
                //onChange={this._onItemsChange}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              ensureUser={true}
              />*/}
            </div>
          </div>
          <Separator></Separator>
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              width: "100%",
            }}
          >
            <PrimaryButton
              text="Submit"
              className="txtclassSubmitBtn"
              onClick={mandatoryvalidation}
              style={{ marginRight: "0.5rem" }}
            />
            <DefaultButton
              text="Cancel"
              onClick={DialogBox}
              className="txtclassCancelBtn"
              styles={{ root: { marginRight: "0.5rem" } }}
            />
          </div>
          <Dialog hidden={Hidedialog} dialogContentProps={dialogContentProps}>
            <DialogFooter>
              <PrimaryButton onClick={DialogBox} text="Ok" />
            </DialogFooter>
          </Dialog>
        </div>
      </div>
    </div>
  );

  function handlechange(e): void {
    var name: string = e.target.attributes.name.value;
    var value: string = e.target.value;
    if (value) {
      Validation[0][name] = "";
      Submitvalues[0][name] = value;
    } else {
      Submitvalues[0][name] = "";
    }

    setSubmitvalues([...Submitvalues]);
    setValidation([...Validation]);
  }

  function UpdateSelectedUsers(item)
  {
    console.log("Called parent function");
    var selctedppls=[];
    item.forEach(async element => {
      await selctedppls.push(element.ID);
    });
    setSelectedpeoples(selctedppls);
  }

  function isEmail(email) {
    var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
    return regex.test(email);
  }

  /*----------------------------------------mandatoryvalidation--------------------------------------*/
  function mandatoryvalidation(): void {
    var isAllFieldsFilled: boolean = true;

    if (!Submitvalues[0].DeviceCount) {
      Validation[0].DeviceCount = "Please Enter DeviceCount";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].PointCount) {
      Validation[0].PointCount = "Please Enter PointCount";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].Drivers) {
      Validation[0].Drivers = "Please Enter Drivers";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].BEName) {
      Validation[0].BEName = "Please Enter site name";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].BENumber) {
      Validation[0].BENumber = "Please Enter BENumber";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].SpecialConsiderations) {
      Validation[0].SpecialConsiderations =
        "Please Enter SpecialConsiderations";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].ManagerName) {
      Validation[0].ManagerName = "Please Enter Name";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].ManagerEmail || !isEmail(Submitvalues[0].ManagerEmail)) {
      Validation[0].ManagerEmail = "Please Enter Email";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].ManagerPhoneNumber) {
      Validation[0].ManagerPhoneNumber = "Please Enter PhoneNumber";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].VendorManagerName) {
      Validation[0].VendorManagerName = "Please Enter Name";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].VendorManagerEmail||!isEmail(Submitvalues[0].VendorManagerEmail)) {
      Validation[0].VendorManagerEmail = "Please Enter Email";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].VendorManagerPhoneNumber) {
      Validation[0].VendorManagerPhoneNumber = "Please Enter PhoneNumber";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].ShippingCountName) {
      Validation[0].ShippingCountName = "Please Enter ShippingCountName";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].ShippingAddress) {
      Validation[0].ShippingAddress = "Please Enter ShippingAddress";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].AdditionalInformation) {
      Validation[0].AdditionalInformation =
        "Please Enter AdditionalInformation";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].Description) {
      Validation[0].Description = "Please Enter Description";
      isAllFieldsFilled = false;
    }
    /*else if(getSelectedUsers.length==0)
    {
      Validation[0].UserDetailsId="Please Enter UserDetails";
      isAllFieldsFilled=false;
    }*/

    setValidation([...Validation]);

    Submit(isAllFieldsFilled);
  }

  async function Submit(allvaluesfilled): Promise<void> {
    if (allvaluesfilled) {
      await setColumn(true);
      var requestdata: formvalues = {
        DeviceCount: Submitvalues[0].DeviceCount,
        PointCount: Submitvalues[0].PointCount,
        Drivers: Submitvalues[0].Drivers,
        SpecialConsiderations: Submitvalues[0].SpecialConsiderations,
        BEName: Submitvalues[0].BEName,
        BENumber: Submitvalues[0].BENumber,
        ManagerName: Submitvalues[0].ManagerName,
        ManagerPhoneNumber: Submitvalues[0].ManagerPhoneNumber,
        ManagerEmail: Submitvalues[0].ManagerEmail,
        VendorManagerName: Submitvalues[0].VendorManagerName,
        VendorManagerPhoneNumber: Submitvalues[0].VendorManagerPhoneNumber,
        VendorManagerEmail: Submitvalues[0].VendorManagerEmail,
        ShippingCountName: Submitvalues[0].ShippingCountName,
        ShippingAddress: Submitvalues[0].ShippingAddress,
        AdditionalInformation: Submitvalues[0].AdditionalInformation,
        Description: Submitvalues[0].Description,
        OrderNo: NextOrderID,
        Category: Category.toString(),
        //UserDetailsId: { results: getSelectedUsers },
        UserDetailsId: { results: Selectedpeoples },
      };
      await props.spcontext.lists
        .getByTitle("WFQuoteRequestList")
        .items.add(requestdata)
        .then(async function (data): Promise<void> {
          console.log(data);
          setColumn(false);
          setHidedialog(false);
        })
        .catch(function (error): void {
          alert(error);
        });
    }
  }
  function DialogBox(): void {
    location.href = siteURL + "/SitePages/WFRequestDashboard.aspx";

    /*var element = document.getElementById('adminFormRequest');
    var opt = {
      margin:       1,
      filename:     'myfile.pdf',
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2 },
      jsPDF:        { unit: 'in', format: 'letter', orientation: 'Landscape' }
    }
    html2pdf().from(element).set(opt).save();
    
  }
  async function getUserID(event): Promise<void> {
    for (let i = 0; i < event.length; i++) {
      getSelectedUsers = [];
      await props.spcontext.siteUsers
        .getByEmail(event[i].secondaryText)
        .get()
        .then(async function (result): Promise<void> {
          if (result.Id) getSelectedUsers.push(result.Id);
        })
        .catch(function (error): void {
          alert(error);
        });
    }*/
  }

  function autoIncrementCustomId(lastRecordId) {
    let increasedNum = Number(lastRecordId.replace("WF-", "")) + 1;
    let kmsStr = lastRecordId.substr(0, 3);

    kmsStr = kmsStr + increasedNum.toString();
    console.log(kmsStr);
    NextOrderID = kmsStr;
  }

  async function getLastID() {
    await props.spcontext.lists
      .getByTitle("WFQuoteRequestList")
      .items.select("ID", "OrderNo")
      .top(1)
      .orderBy("ID", false)
      .get()
      .then(function (data) {
        if (data.length > 0) {
          autoIncrementCustomId(data[0].OrderNo);
        } else {
          autoIncrementCustomId("WF-0");
        }
      })
      .catch(function (error) {});
  }
}
export { RequestNewQuoteAdmin };
