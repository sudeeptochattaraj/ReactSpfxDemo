import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IStackTokens, Stack } from '@fluentui/react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
//import { DropDownButtonComponent, ItemModel } from '@syncfusion/ej2-react-splitbuttons';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton, IconButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { FontWeights, getTheme, mergeStyleSets } from 'office-ui-fabric-react';
import { IIconProps } from '@fluentui/react';
import { ActionButton } from '@fluentui/react/lib/Button';
let options: IDropdownOption[];
let Cityoptions: IDropdownOption[];
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};
const stackTokens: IStackTokens = { childrenGap: 20 };
export interface IStates {
  Items: any;
  ID: any;
  EmployeeName: any;
  EmployeeNameId: any;
  HireDate: any;
  JobDescription: any;
  HTML: any;
  options: any;
  selectedOption: '';
  singleValueDropdown: string;
  Cityoptions: any;
  CityNameId: any;
  StateNameId: any;
  cityName: any;
  StateName: any;
  showModal: boolean;
  disabled?: boolean;
  checked?: boolean;
  createbutton?: boolean;
  Updatebutton?: boolean;
  Deletebutton?: boolean;

}
export default class CRUDReact extends React.Component<IHelloWorldProps, IStates> {
  constructor(props) {
    super(props);
    this.state = {
      Items: [],
      EmployeeName: "",
      EmployeeNameId: 0,
      ID: 0,
      HireDate: null,
      JobDescription: "",
      HTML: [],
      options: [],
      selectedOption: "",
      singleValueDropdown: "",
      Cityoptions: [],
      CityNameId: 0,
      StateNameId: 0,
      cityName: "",
      StateName: "",
      showModal: false,
      disabled: false,
      checked: false,
      createbutton: true,
      Updatebutton: true,
      Deletebutton: true
      

    };
  }

  public async componentDidMount() {
    await this.fetchData();
    this.Dropdownbind();
  }

  //hi sudeepto
  public async Dropdownbind() {

    options = [];
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists.getByTitle("State").items.select("*").get();
    console.log(items);
    items.forEach(c => {

      options.push({
        key: c.ID,
        text: c.Title
      });
    });

    this.setState({ options: options });

  }
  async DropCitydownbind(event) {
    debugger
    Cityoptions = [];
    var newValue = event.key;
    var NewText = event.text;
    let web = Web(this.props.webURL);
    var newValue = event.key;
    const items: any[] = await web.lists.getByTitle("CityDetails").items.select("*").get();
    console.log(items);
    var FilterData = items.filter(el => el.StateNameId == newValue)
    FilterData.forEach(c => {
      Cityoptions.push({
        key: c.ID,
        text: c.Title
      });
    });

    this.setState({ Cityoptions: Cityoptions, StateNameId: newValue, StateName: NewText });

  }
  async DropCitydownbindfor(event) {
    debugger
    var newValue = event.key;
    var NewText = event.text;
    this.setState({ CityNameId: newValue, cityName: NewText });
  }
  public async fetchData() {
    debugger
    let web = Web(this.props.webURL);
    let items: any[] = await web.lists.getByTitle("EmployeeDetails").items.select("*", "EmployeeName/Title").expand("EmployeeName/ID").get();
    this.setState({ Items: items });
    let html = await this.getHTML(items);
    this.setState({ HTML: html });
  }
  public findData = (id): void => {
    debugger
    var itemID = id;
    var allitems = this.state.Items;
    var allitemsLength = allitems.length;
    if (allitemsLength > 0) {
      for (var i = 0; i < allitemsLength; i++) {
        if (itemID == allitems[i].Id) {
          this.setState({
            ID: itemID,
            EmployeeName: allitems[i].EmployeeName.Title,
            EmployeeNameId: allitems[i].EmployeeNameId,
            HireDate: new Date(allitems[i].HireDate),
            JobDescription: allitems[i].Job_x0020_Description,
            CityNameId: allitems[i].CityNameId,
            StateNameId: allitems[i].StateNameId,
            cityName: allitems[i].CityName,
            StateName: allitems[i].StateName,
          });
        }
      }
    }
    this.showModal(2);
  }
  public onSelect(): void {
    alert("Select event is triggered");
  }
  public onBeforeOpen(): void {
    alert("beforeOpen event is triggered");
  }
  private showModal = (a): void => {


    this.setState({ showModal: true });
    if (a == 1) {
      this.setState({ createbutton: false });
    }
    else if (a == 2) {
      this.setState({ Updatebutton: false });
      this.setState({ Deletebutton: false });
      this.setState({ createbutton: true });
    }
  };

  private closeModal = (): void => {
    this.setState({ showModal: false });
  };



  public async getHTML(items) {

    var tabledata = <table className={styles.table}>
      <thead>
        <tr>
          <th>Employee Name</th>
          <th>Hire Date</th>
          <th>Job Description</th>
          <th>State</th>
          <th>City</th>
        </tr>
      </thead>
      <tbody>
        {items && items.map((item, i) => {
          return [
            <tr key={i} onClick={() => this.findData(item.ID)}>
              <td>{item.EmployeeName.Title}</td>
              <td>{FormatDate(item.HireDate)}</td>
              <td>{item.Job_x0020_Description}</td>
              <td>{item.StateName}</td>
              <td>{item.CityName}</td>
            </tr>
          ];
        })}
      </tbody>

    </table>;
    return await tabledata;
  }

  public _getPeoplePickerItems = async (items: any[]) => {

    if (items.length > 0) {

      this.setState({ EmployeeName: items[0].text });
      this.setState({ EmployeeNameId: items[0].id });
    }
    else {
      this.setState({ EmployeeNameId: "" });
      this.setState({ EmployeeName: "" });
    }
  }
  public onchange = async (value, stateValue) => {
    this.setState({ JobDescription: stateValue });
    console.log(stateValue);
  }

  private async SaveData() {
    debugger
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("EmployeeDetails").items.add({

      EmployeeNameId: this.state.EmployeeNameId,
      HireDate: new Date(this.state.HireDate),
      Job_x0020_Description: this.state.JobDescription,
      StateNameId: this.state.StateNameId,
      CityNameId: this.state.CityNameId,
      StateName: this.state.StateName,
      CityName: this.state.cityName
    })
    debugger
    alert("Created Successfully");
    this.closeModal();
    this.fetchData();
  }
  private async UpdateData() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("EmployeeDetails").items.getById(this.state.ID).update({

      EmployeeNameId: this.state.EmployeeNameId,
      HireDate: new Date(this.state.HireDate),
      Job_x0020_Description: this.state.JobDescription,
      StateNameId: this.state.StateNameId,
      CityNameId: this.state.CityNameId,
      StateName: this.state.StateName,
      CityName: this.state.cityName

    }).then(i => {
      console.log(i);
    });
    alert("Updated Successfully");
    this.closeModal();
    this.fetchData();
  }
  private async DeleteData() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("EmployeeDetails").items.getById(this.state.ID).delete()
      .then(i => {
        console.log(i);
      });
    alert("Deleted Successfully");
    this.closeModal();
    this.fetchData();
  }
  private titleId: string = getId('title');
  private subtitleId: string = getId('subText');
  public render(): React.ReactElement<IHelloWorldProps> {

    const { selectedOption } = this.state;
    const value = selectedOption;
    const addFriendIcon: IIconProps = { iconName: 'AddFriend' };
    let i = 1;
    return (

      <div>
        <h1>Employee Details</h1>

        <div className={styles.btngroup}>
          <div>
            <ActionButton iconProps={addFriendIcon} allowDisabledFocus onClick={() => this.showModal(i)}>
              Add New Employee
    </ActionButton>

          </div>

        </div>
        {this.state.HTML}
        <Modal
          titleAriaId={this.titleId}
          subtitleAriaId={this.subtitleId}
          isOpen={this.state.showModal}
          onDismiss={this.closeModal}
          isClickableOutsideFocusTrap={true}
          isBlocking={true}
          containerClassName={contentStyles.container}>

          <div className={contentStyles.container}>
            <h1><span id={this.titleId}>Enter Employee Details</span></h1>
          </div>
          <div className={contentStyles.body}>
            <form>
              <div>
                <Label>Employee Name </Label>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  showtooltip={true}
                  isRequired={true}
                  selectedItems={this._getPeoplePickerItems}
                  defaultSelectedUsers={[this.state.EmployeeName ? this.state.EmployeeName : ""]}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                />
              </div>
              <div>
                <Label>Hire Date</Label>
                <DatePicker maxDate={new Date()} allowTextInput={false} strings={DatePickerStrings} value={this.state.HireDate} onSelectDate={(e) => { this.setState({ HireDate: e }); }} ariaLabel="Select a date" formatDate={FormatDate} />
              </div>
              <div>

                <Label>Job Description</Label>
                <TextField value={this.state.JobDescription} multiline onChange={this.onchange} />
              </div>
              <Stack tokens={stackTokens}>
                <Dropdown
                  placeholder="Select a State"
                  label="State"
                  options={options}
                  styles={dropdownStyles}
                  selectedKey={this.state.StateNameId ? this.state.StateNameId : undefined}
                  onChanged={this.DropCitydownbind.bind(this)}
                />
                <Dropdown
                  placeholder="Select a city"
                  label="City"
                  options={Cityoptions}
                  styles={dropdownStyles}
                  selectedKey={this.state.CityNameId ? this.state.CityNameId : undefined}
                  onChanged={this.DropCitydownbindfor.bind(this)}
                />
                <div className={styles.btngroup}>
                  <div><PrimaryButton text="Create" disabled={this.state.createbutton} onClick={() => this.SaveData()} /></div>
                  <div><PrimaryButton text="Update" disabled={this.state.Updatebutton} onClick={() => this.UpdateData()} /></div>
                  <div><PrimaryButton text="Delete" disabled={this.state.Deletebutton} onClick={() => this.DeleteData()} /></div>
                  <div><DefaultButton onClick={this.closeModal} text="Close" /></div>
                </div>

              </Stack>
            </form>
          </div>

        </Modal>
      </div>
    );
  }
}
export const DatePickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  invalidInputErrorMessage: 'Invalid date format.'
};
export const FormatDate = (date): string => {
  console.log(date);
  var date1 = new Date(date);
  var year = date1.getFullYear();
  var month = (1 + date1.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;
  var day = date1.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  return month + '/' + day + '/' + year;
};

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const toggleStyles = { root: { marginBottom: '20px' } };
const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
