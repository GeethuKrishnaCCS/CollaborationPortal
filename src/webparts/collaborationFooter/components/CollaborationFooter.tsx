import * as React from 'react';
import styles from './CollaborationFooter.module.scss';
import type { ICollaborationFooterProps, ICollaborationFooterState } from '../interfaces/ICollaborationFooterProps';
import { DefaultButton, FontWeights, IIconProps, IconButton, Modal, PrimaryButton, TextField, getTheme, mergeStyleSets, mergeStyles } from '@fluentui/react';
import { CollaborationFooterService } from '../services';


export default class CollaborationFooter extends React.Component<ICollaborationFooterProps,  ICollaborationFooterState, {}> {
 private _Service: CollaborationFooterService;
  public constructor(props: ICollaborationFooterProps) {
    super(props);
    this.state = {
      currentUserId: null,
      isAdmin: false,
      logoUrl: "",
      logolink: "",
      linkedinLogoUrl: "",
      footerdata: [],
      linkedinLogolink: "",
      openEditModal: false,
      footerdataitem: [],
      openEditItemModal: false,
      title: "",
      itemLink: "",
      itemId: null,
      openAddModal: false,
      itemLogo: null
    }
    this._Service = new CollaborationFooterService(this.props.context);
    this.checkuserAdmin = this.checkuserAdmin.bind(this);
    this.editFooter = this.editFooter.bind(this);
    this.editFooterModalClose = this.editFooterModalClose.bind(this);
    this.editFooterItemModalClose = this.editFooterItemModalClose.bind(this);
    this.editFooterItem = this.editFooterItem.bind(this);
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onLinkChange = this.onLinkChange.bind(this);
    this.onUpdateItem = this.onUpdateItem.bind(this);
    this.addnewitemModal = this.addnewitemModal.bind(this);
    this.addnewitemModalClose = this.addnewitemModalClose.bind(this);
    this.onSubmitItem = this.onSubmitItem.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.checkuserAdmin();
    const footerdataArray: any[] = []
    const footerdata = await this._Service.getListItems(this.props.siteUrl, this.props.settingsList)
    for (let i = 0; i < footerdata.length; i++) {
      if (footerdata[i].Title !== "Company Logo" && footerdata[i].Title !== "LINKED IN") {
        const footerda = {
          Title: footerdata[i].Title,
          Link: footerdata[i].Link,
        };
        footerdataArray.push(footerda);
      }
      this.setState({
        footerdata: footerdataArray,
      });
    }
    const CompanyLogo = footerdata.filter((item: any) => item.Title === "Company Logo");
    if (CompanyLogo.length > 0) {
      this.setState({ logoUrl: CompanyLogo[0].Logo.Url, logolink: CompanyLogo[0].Link });
    }
    const LinkedinLogo = footerdata.filter((item: any) => item.Title === "LINKED IN");
    if (LinkedinLogo.length > 0) {
      this.setState({ linkedinLogoUrl: LinkedinLogo[0].Logo.Url, linkedinLogolink: LinkedinLogo[0].Link });
    }

  }
  public async checkuserAdmin() {
    const currentuser = await this._Service.getCurrentUser();
    if (currentuser) {
      const admindata = await this._Service.getAdminListItems(this.props.siteUrl, this.props.adminList)
      const userAdmin = admindata.filter((item: any) => item.Admin.ID === currentuser.Id);
      if (userAdmin.length > 0) {
        this.setState({ isAdmin: true });
      }
    }

  }
  public async editFooter() {
    let LogoUrl: string = "";
    let footerdataitems: any[] = [];
    const footerdata = await this._Service.getListItems(this.props.siteUrl, this.props.settingsList)
    for (let i = 0; i < footerdata.length; i++) {
      if (footerdata[i].Logo !== null) {
        LogoUrl = footerdata[0].Logo.Url;
      }
      const footerdataitem = {
        Title: footerdata[i].Title,
        Link: footerdata[i].Link,
        Logo: footerdata[i].Logo,
        LogoUrl: LogoUrl,
        Id: footerdata[i].ID

      };
      footerdataitems.push(footerdataitem)
    }
    this.setState({
      footerdataitem: footerdataitems
    });
    this.setState({
      openEditModal: true
    });
  }
  public editFooterModalClose() {
    this.setState({
      openEditModal: false,
      openEditItemModal: false,
      openAddModal: false
    });
  }
  public editFooterItemModalClose() {
    this.setState({
      openEditItemModal: false, openEditModal: true
    });
  }
  public editFooterItem(Id: any) {
    const itemDetails = this.state.footerdataitem.filter((item: any) => item.Id === Id);
    if (itemDetails.length > 0) {
      this.setState({
        title: itemDetails[0].Title,
        itemLink: itemDetails[0].Link,
        itemId: itemDetails[0].Id,
        itemLogo: itemDetails[0].Logo
      });
    }
    this.setState({
      openEditItemModal: true, openEditModal: false
    });

  }
  public addnewitemModal() {
    this.setState({
      title: "",
      itemLink: "", openAddModal: true, openEditModal: false, openEditItemModal: false
    });
  }
  public addnewitemModalClose() {
    this.setState({
      title: "",
      itemLink: "",openAddModal: false, openEditModal: true, openEditItemModal: false
    });
  }

  //Title Change
  private onTitleChange = (ev: React.FormEvent<HTMLInputElement>, Title: string): void => {
    this.setState({
      title: Title,
    });
  }
  //Link Change
  private onLinkChange = (ev: React.FormEvent<HTMLInputElement>, itemLink: string): void => {
    this.setState({
      itemLink: itemLink,
    });
  }
  public async onUpdateItem() {
    const datafooter = {
      Title: this.state.title,
      Link: this.state.itemLink
    }
    const updatelist = await this._Service.updateItem(this.props.siteUrl, this.props.settingsList, datafooter, this.state.itemId)
    if (updatelist) {
      this.setState({ openEditItemModal: false, openEditModal: false,openAddModal: false });
    }
  }
  public async onSubmitItem() {
    if (this.state.title !== "" || this.state.itemLink !== "") {
      const datafooter = {
        Title: this.state.title,
        Link: this.state.itemLink
      }
      const updatelist = await this._Service.createNewItem(this.props.siteUrl, this.props.settingsList, datafooter)
      if (updatelist) {
        this.setState({ openAddModal: false, openEditModal: false,openEditItemModal:false });
      }
    }
  }


  public render(): React.ReactElement<ICollaborationFooterProps> {
    const theme = getTheme();
    const borderlessButtonStyle = mergeStyles({
      border: 'none',
      padding: 0,
      backgroundColor: 'transparent',
      // color: '#3797e4',
      selectors: {
        ':hover': {
          backgroundColor: 'transparent',
          // color: '#ffffff'
        },
        ':active': {
          backgroundColor: 'transparent',
          // color: '#3797e4'
        },
      },
    });
    const contentStyles = mergeStyleSets({
      container: {

        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        borderRadius: '1rem',
        padding: '0.5rem'


      },
      header: [
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
      heading: {
        color: theme.palette.neutralPrimary,
        fontWeight: FontWeights.semibold,
        fontSize: 'inherit',
        margin: '0',
      },
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
    const CancelIcon: IIconProps = { iconName: 'Cancel' };
    const MoreIcon: IIconProps = { iconName: 'More' }
    const AddIcon: IIconProps = { iconName: 'CircleAddition' }
    const iconButtonStyles = {
      root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '1px',
      },
      rootHovered: {
        color: theme.palette.neutralDark,
      },
    };
    const AddiconButtonStyles = {
      root: {
        color: theme.palette.neutralPrimary,
        background: 'white',
        padding: '0px',
        margin: '0px',
        width: '16px',
        height: '16px',
      },
      rootHovered: {
        color: theme.palette.neutralDark,
      },
    };

    return (
      <section className={`${styles.collaborationFooter}`}>
        
         <div className={styles.msGrid} dir="ltr" >
          <div className={styles.msRow}>
            <div className={styles.msCol} >
              <div className={styles.footerbg}>
                <div className={styles.footercntnt}>
                  <div className={styles.companylogo}><a href={this.state.logolink} target="_blank"><img src={this.state.logoUrl} /></a></div>
                  <div className={styles.itemstyle}>{this.state.footerdata.map((item: any) => (
                    <a href={item.Link} target="_blank">{item.Title}</a>
                  ))}</div>
                  <div className={styles.footerright}>
                    <div className={styles.linkedinLogo}><a href={this.state.linkedinLogolink} target="_blank"><img src={this.state.linkedinLogoUrl} /></a></div>
                    {this.state.isAdmin && <div><DefaultButton
                      className={`${borderlessButtonStyle} ${styles.editftrbutton}`}
                      onClick={this.editFooter}
                      text="Edit"
                    /></div>}
                  </div> 
                </div>
                {/* Edit Footer */}
                <Modal
                  isOpen={this.state.openEditModal}
                  isModeless={false}
                  containerClassName={contentStyles.container}>
                  <div style={{ padding: "18px" }}>
                    <div style={{ display: "flex" }}>
                      <span style={{ textAlign: "center", display: "flex", justifyContent: "center", flexGrow: "1", fontSize: "20px", fontFamily: 'sans-serif', fontWeight: "400" }}><b>Edit Footer</b></span>
                      <IconButton
                        iconProps={CancelIcon}
                        ariaLabel="Close popup modal"
                        onClick={this.editFooterModalClose}
                        styles={iconButtonStyles} />
                    </div>
                    <div>
                      <table>
                        {this.state.footerdataitem.map((item: any) => (
                          <tr>
                            <td>{item.Title}</td>
                            <td> <IconButton
                              styles={iconButtonStyles}
                              iconProps={MoreIcon}
                              ariaLabel="More options"
                              onClick={() => this.editFooterItem(item.Id)}
                            /></td>
                          </tr>
                        ))}
                      </table>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", position: "relative" }}>
                        <hr style={{ position: "absolute", width: "100%" }} /><IconButton
                          iconProps={AddIcon}
                          ariaLabel="Add NewItem modal"
                          onClick={this.addnewitemModal}
                          styles={AddiconButtonStyles} /></div>
                    </div>
                  </div>

                </Modal>
                {/* Edit Item Footer */}
                <Modal
                  isOpen={this.state.openEditItemModal}
                  isModeless={false}
                  containerClassName={contentStyles.container}>
                  <div style={{ padding: "18px" }}>
                    <div style={{ display: "flex" }}>
                      <span style={{ textAlign: "center", display: "flex", justifyContent: "center", flexGrow: "1", width: "450px", fontSize: "20px", fontFamily: 'sans-serif', fontWeight: "400" }}><b>Edit Footer Item</b></span>
                      <IconButton
                        iconProps={CancelIcon}
                        ariaLabel="Close popup modal"
                        onClick={this.editFooterItemModalClose}
                        styles={iconButtonStyles} />
                    </div>
                    <div>
                      <table style={{ width: "100%" }}>
                        {this.state.itemLogo === null && <tr>
                          <td><TextField
                            label='Title'
                            onChange={this.onTitleChange}
                            value={this.state.title} />
                          </td>
                        </tr>}
                        <tr>
                          <td><TextField
                            label='Link'
                            onChange={this.onLinkChange}
                            value={this.state.itemLink} />
                          </td>
                        </tr>
                      </table>
                      <PrimaryButton style={{ float: "right", marginTop: "7px", marginBottom: "9px" }} id="b2" onClick={this.onUpdateItem} >Save</PrimaryButton >
                    </div>
                  </div>

                </Modal>
                {/* Add Item Footer */}
                <Modal
                  isOpen={this.state.openAddModal}
                  isModeless={false}
                  containerClassName={contentStyles.container}>
                  <div style={{ padding: "18px" }}>
                    <div style={{ display: "flex" }}>
                      <span style={{ textAlign: "center", display: "flex", justifyContent: "center", flexGrow: "1", width: "450px", fontSize: "20px", fontFamily: 'sans-serif', fontWeight: "400" }}><b>Add Footer Item</b></span>
                      <IconButton
                        iconProps={CancelIcon}
                        ariaLabel="Close popup modal"
                        onClick={this.addnewitemModalClose}
                        styles={iconButtonStyles} />
                    </div>
                    <div>
                      <table style={{ width: "100%" }}>
                        <tr>
                          <td><TextField
                            label='Title'
                            onChange={this.onTitleChange}
                            value={this.state.title} />
                          </td>
                        </tr>
                        <tr>
                          <td><TextField
                            label='Link'
                            onChange={this.onLinkChange}
                            value={this.state.itemLink} />
                          </td>
                        </tr>
                      </table>
                      <PrimaryButton style={{ float: "right", marginTop: "7px", marginBottom: "9px" }} id="b2" onClick={this.onSubmitItem} >Submit</PrimaryButton >
                    </div>
                  </div>

                </Modal>
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
