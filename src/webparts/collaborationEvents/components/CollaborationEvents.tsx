import * as React from 'react';
import styles from './CollaborationEvents.module.scss';
import { ICollaborationEventsProps, ICollaborationEventsState } from '../interfaces';
import { ActionButton, Checkbox, DatePicker, DefaultButton, Dropdown, FocusTrapZone, Icon, IDropdownOption, Label, Layer, Overlay, Popup, PrimaryButton, TextField, TimePicker } from '@fluentui/react';

import { CollaborationEventsService } from '../services/CollaborationEventsService';
import FilmstripView from '../views/FilmstripView';


export default class EventsTemplate extends React.Component<ICollaborationEventsProps, ICollaborationEventsState, {}> {

  private _service: any;

  public constructor(props: ICollaborationEventsProps) {
    super(props);

    this._service = new CollaborationEventsService(this.props.context, this.props.context.pageContext.web.serverRelativeUrl);
    this.state = {
      EventCategory: [],
      ListDetails: [],
      EventTitle: "",
      Eventlocation: "",
      EventDescription: "",
      EventLink: "",
      selectedCategory: { key: "", text: "" },

      StartDate: new Date(),  // Default to current date
      StartTime: new Date(),  // Default to current time

      EndDate: new Date(),
      EndTime: new Date(),

      AlldayNeeded: false,
      EnableEventTemplate: true,
      isPopupVisible: false,

      selectedImage: undefined,
      EventImage: "" // Initialize EventImage
    }

    this.getchoice = this.getchoice.bind(this);
    this.getEventList = this.getEventList.bind(this);
    this.onChangeEventTitle = this.onChangeEventTitle.bind(this);
    this.onChangeLocation = this.onChangeLocation.bind(this);
    this.onChangeLink = this.onChangeLink.bind(this);
    this.onChangeDescription = this.onChangeDescription.bind(this);
    this.getOnchangeCategory = this.getOnchangeCategory.bind(this);
    this.onSelectDate = this.onSelectDate.bind(this);
    this.onSelectTime = this.onSelectTime.bind(this);
    this.onAlldayNeededChange = this.onAlldayNeededChange.bind(this);

    this.onSubmitClick = this.onSubmitClick.bind(this);

  }

  public async componentDidMount() {
    await this.getchoice();
    await this.getEventList();

  }

  public async getchoice() {
    const url: string = `${this.props.context.pageContext.web.serverRelativeUrl}/Lists/EventsList`
    const getdeliveryModechoice = await this._service.getChoiceListItems(url, "Category");
    console.log('getdeliveryModechoice: ', getdeliveryModechoice);

    const deliveryMode: { key: string, text: string }[] = [];
    getdeliveryModechoice.Choices.map((item: string, index: number) => {
      deliveryMode.push({ key: index.toString(), text: item });
    });
    this.setState({ EventCategory: deliveryMode });
  }
  public onChangeEventTitle(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, eventTitle: string) {

    this.setState({ EventTitle: eventTitle });

  }

  public onChangeLocation(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, eventLocation: string) {

    this.setState({ Eventlocation: eventLocation });

  }

  public onChangeDescription(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, eventDescription: string) {

    this.setState({ EventDescription: eventDescription });

  }

  public onChangeLink(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, eventLink?: string) {
    this.setState({ EventLink: eventLink || "" });
  }

  public getOnchangeCategory(event: React.FormEvent<HTMLDivElement>, category: IDropdownOption): void {
    this.setState({ selectedCategory: category });
  }



  public onSelectDate = (date: Date): void => {
    this.setState({ StartDate: date });
  };

  public onSelectTime = (time: Date): void => {
    this.setState({ StartTime: time });
  };


  //For END DATE AND TIME

  public onSelectEndDate = (date: Date): void => {
    this.setState({ EndDate: date });
  };

  public onSelectEndTime = (time: Date): void => {
    this.setState({ EndTime: time });
  };
  //CLOSE THE FUNCTION  


  public onAlldayNeededChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ AlldayNeeded: isChecked });
  }

  public async onSubmitClick(): Promise<void> {
    const StartDate: Date = this.state.StartDate instanceof Date
      ? this.state.StartDate
      : new Date(this.state.StartDate);

    const StartTime: Date = this.state.StartTime instanceof Date
      ? this.state.StartTime
      : new Date(this.state.StartTime);

    const EndDate: Date = this.state.EndDate instanceof Date
      ? this.state.EndDate
      : new Date(this.state.EndDate);

    const EndTime: Date = this.state.EndTime instanceof Date
      ? this.state.EndTime
      : new Date(this.state.EndTime);

    const EndDateTimeUTC = new Date(
      EndDate.getFullYear(), EndDate.getMonth(), EndDate.getDate(),
      EndTime.getHours(), EndTime.getMinutes()
    );

    // Combine Date and Time
    StartDate.setHours(StartTime.getHours(), StartTime.getMinutes());
    EndDate.setHours(EndDate.getHours(), EndTime.getMinutes());

    const dataItem = {
      Title: this.state.EventTitle,
      Location: this.state.Eventlocation,
      Description: this.state.EventDescription,
      BannerURL: {
        Description: "Event Link",
        Url: this.state.EventLink
      },
      Category: this.state.selectedCategory.text,
      StartTime: StartDate.toISOString(),
      EndDate: EndDateTimeUTC.toISOString(),
      AllDay: this.state.AlldayNeeded,
      EventImage: "" // Initialize as empty string
    };

    try {
      const attachment = this.state.selectedImage;
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      const newItem = await this._service.addItemRequestForm(dataItem, "EventsList", url);

      if (attachment) {
        const attachmentResponse = await this._service.addAttachments("EventsList", newItem.data.ID, attachment.name, attachment);
        const attachmentUrl = attachmentResponse.data.ServerRelativeUrl;
        await this._service.updateItemRequestForm({ EventImage: attachmentUrl }, "EventsList", newItem.data.ID, url);
      }

      console.log("Successfully added item:", newItem);
      alert("Successfully Completed");
    } catch (error) {
      console.error("Error submitting request:", error);
      alert("Error submitting request");
    }

    this.setState({ isPopupVisible: false });
  }

  public async getEventList() {
    const url: string = `${this.props.context.pageContext.web.serverRelativeUrl}/Lists/EventsList`
    const getEventList = await this._service.getListItems(url);
    const listDetailsWithImages = getEventList.map((item: any) => ({
      ...item,
      EventImage: item.EventImage ? `${url}/_layouts/15/getpreview.ashx?path=${item.EventImage}` : ""
    }));
    this.setState({ ListDetails: listDetailsWithImages });
    console.log('ListDetails: ', this.state.ListDetails);
  }

  public hidePopup = () => {
    this.setState({ isPopupVisible: false });
  };


  public render(): React.ReactElement<ICollaborationEventsProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.collaborationEvents} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2 style={{ color: this.props.headingfontcolor }}>{this.props.description}</h2>

          <div className={styles.layout}>
            <div className={styles.card}>

              <div className={styles.toprightbutton}>
                <ActionButton
                  iconProps={{ iconName: "Add" }}
                  onClick={() => this.setState({ isPopupVisible: true })}                >
                  Add Event
                </ActionButton>
              </div>


              {/* {this.props.Cardlayout === "Filmstrip" && (
                <>
                  {this.state.ListDetails.length > 0 && (
                    <div className={styles.StackStyle}>
                      <div className={styles.StackStyleContainer}>
                        <div className={styles.BirthdaySlider}>
                          {this.state.ListDetails
                            .filter((event: any) =>
                              !this.props.location || event.Location === this.props.location
                            )
                            .map((event: any) => {
                              const eventImageUrl = event.EventImage ? event.EventImage.split('?path=')[1] : "https://via.placeholder.com/150";
                              return (
                                <div className={styles.FilmstripCard} key={event.Id}>

                                  <div className={styles.imageContainer}>
                                    <img className={`${styles.profileImage} ${this.props.StylesForImages === "Square"
                                      ? styles.squareImageCard
                                      : this.props.StylesForImages === "Circle"
                                        ? styles.circleImageCard
                                        : this.props.StylesForImages === "Rectangle"
                                          ? styles.rectangleImageCard
                                          : ""
                                      }`}
                                      src={event.EventImage ? `https://ccsdev01.sharepoint.com${eventImageUrl}` : require('../assets/DefaultImage.png')}
                                      alt={event.Title}
                                    // className={styles.cardImage}
                                    />

                                    <div className={styles.date}>
                                      {new Date(event.StartTime).toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                                    </div>

                                  </div>


                                  <div className={styles.cardContent}>
                                    <span className={styles.category}>{event.Category}</span>
                                    <h3 className={styles.title}>{event.Title}</h3>
                                    <p className={styles.dateTime}>
                                      {new Date(event.StartTime).toLocaleString("en-US", {
                                        weekday: "short",
                                        month: "short",
                                        day: "numeric",
                                        hour: "2-digit",
                                        minute: "2-digit",
                                      })}
                                    </p>

                                    <p>{event.Location}</p>

                                    <div className={styles.actionButtons}>

                                      <img
                                        src={require('../assets/RSVP.png')}
                                        className={styles.rsvpButton}
                                      />

                                      <img
                                        src={require('../assets/Mail.png')}
                                        className={styles.mailIcon}
                                      />


                                      <img
                                        src={require('../assets/Comment.png')}
                                        className={styles.commentIcon}
                                      />

                                      <img
                                        src={require('../assets/Like.png')}
                                        className={styles.likeIcon}
                                      />

                                    </div>

                                  </div>
                                </div>
                              );
                            })}
                        </div>
                      </div>
                    </div>
                  )}
                  {this.state.ListDetails.length === 0 && (
                    <div className={styles.nobirthday}>
                      <svg width="64" height="64" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M7.937 41.938L15.062 49.063L19.875 46.875L10.125 37.125L7.937 41.938Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M1.999 55.001L10.25 51.25L5.75 46.75L1.999 55.001Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M53.875 37.125L44.125 46.875L48.937 49.063L56.062 41.938L53.875 37.125Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M58.25 46.75L53.75 51.25L62.001 55.001L58.25 46.75Z" fill="#939191" fillOpacity="0.33" />
                      </svg>
                    </div>
                  )}
                </>
              )} */}

              {this.props.Cardlayout === "Filmstrip" && (
                <>
                  {this.state.ListDetails.length > 0 ? (
                    <FilmstripView
                      events={this.state.ListDetails}
                      location={this.props.location}
                      stylesForImages={this.props.StylesForImages}
                    />
                  ) : (
                    <div className={styles.nobirthday}>
                      <svg width="64" height="64" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M7.937 41.938L15.062 49.063L19.875 46.875L10.125 37.125L7.937 41.938Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M1.999 55.001L10.25 51.25L5.75 46.75L1.999 55.001Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M53.875 37.125L44.125 46.875L48.937 49.063L56.062 41.938L53.875 37.125Z" fill="#939191" fillOpacity="0.33" />
                        <path d="M58.25 46.75L53.75 51.25L62.001 55.001L58.25 46.75Z" fill="#939191" fillOpacity="0.33" />
                      </svg>
                    </div>
                  )}
                </>
              )}





              {this.props.Cardlayout === "Compact" && (
                <>
                  {this.state.ListDetails.length > 0 ? (
                    <div className={styles.CompactGrid}>
                      {this.state.ListDetails
                        .filter((event: any) =>
                          this.props.location ? event.Location === this.props.location : true
                        )
                        .map((event: any, index: any) => {
                          const eventImageUrl = event.EventImage ? event.EventImage.split('?path=')[1] : 'defaultImage.png';
                          return (
                            <div className={styles.CompactCard} key={event.Id}>
                              <div className={styles.cardImageCompact}>
                                <img
                                  src={event.EventImage ? `https://ccsdev01.sharepoint.com${eventImageUrl}` : require('../assets/DefaultImage.png')}
                                  alt={event.Title}
                                  className={styles.cardImageCompact}
                                />
                              </div>
                              <div className={styles.cardDetails}>
                                <h3>{event.Title}</h3>
                                <p className={styles.cardRole}>{event.Category}</p>
                                <p className={styles.cardDate}>
                                  {new Date(event.StartTime).toLocaleString('en-US', {
                                    weekday: 'short',
                                    month: 'short',
                                    day: 'numeric',
                                    hour: '2-digit',
                                    minute: '2-digit',
                                  })}<br />
                                </p>
                              </div>
                            </div>
                          );
                        })}
                    </div>
                  ) : (
                    <div className={styles.noData}>
                      <p>No events available</p>
                    </div>
                  )}
                </>
              )}


              {this.props.Cardlayout === "Cards" && (
                <>
                  {this.state.ListDetails.length > 0 ? (
                    <div className={styles.CardsGrid}>
                      {this.state.ListDetails
                        .filter((event: any) =>
                          this.props.location ? event.Location === this.props.location : true
                        )
                        .map((event: any, index: any) => {
                          const eventImageUrl = event.EventImage
                            ? event.EventImage.split('?path=')[1]
                            : 'defaultImage.png';
                          return (
                            <div className={styles.Card} key={event.Id}>
                              <div className={styles.cardDateContainer}>
                                <span className={styles.cardMonth}>
                                  {new Date(event.Created).toLocaleString('en-US', {
                                    month: 'short',
                                  }).toUpperCase()}
                                </span>
                                <span className={styles.cardDay}>
                                  {new Date(event.Created).getDate()}
                                </span>
                              </div>
                              <div className={this.props.StylesForImages === "Circle" ? styles.circleImageCard : this.props.StylesForImages === "Square" ? styles.squareImageCard : styles.cardImageContainer}>
                                <img
                                  src={
                                    event.EventImage
                                      ? `https://ccsdev01.sharepoint.com${eventImageUrl}`
                                      : require('../assets/DefaultImage.png')
                                  }
                                  alt={event.Title}
                                  className={styles.cardImage}
                                />
                              </div>
                              <div className={styles.cardDetails}>
                                <h3>{event.Title}</h3>
                                <p className={styles.cardRole}>{event.Category}</p>
                                <p><strong>Location:</strong> {event.Location}</p>
                                <p><strong>Start Date:</strong> {new Date(event.StartTime).toLocaleString()}</p>
                                <p><strong>End Date:</strong> {new Date(event.EndDate).toLocaleString()}</p>
                                <div className={styles.cardIcons}>
                                  <i className="fas fa-envelope"></i>
                                  <i className="fas fa-search"></i>
                                  <i className="fas fa-heart"></i>
                                </div>
                              </div>
                            </div>
                          );
                        })}

                    </div>
                  ) : (
                    <div className={styles.noData}>
                      <p>No events available</p>
                    </div>
                  )}
                </>
              )}



              {this.props.Cardlayout === "List" && (
                <div className={styles.listContainer}>
                  {/* <div className={styles.listItems}> */}
                  <div className={this.props.StylesForCards === "Circle" ? styles.listItems : this.props.StylesForCards === "Square" ? styles.squareImage : styles.listItem}>
                    {this.state.ListDetails.map((event: any) => (
                      <div className={styles.listItem} key={event.Id}>
                        <div className={this.props.StylesForImages === "Circle" ? styles.circleImage : this.props.StylesForImages === "Square" ? styles.squareImage : styles.cardImageContainerList}>
                          <img
                            src={
                              event.EventImage
                                ? `https://ccsdev01.sharepoint.com${event.EventImage.split('?path=')[1]}`
                                : require('../assets/DefaultImage.png')
                            }
                            alt={event.Title}
                            className={styles.cardImage}
                          />
                        </div>
                        <div className={styles.details}>
                          <span className={styles.name}>{event.Title}</span>
                          <p><strong>Location:</strong> {event.Location}</p>
                          <p><strong>Start Date:</strong> {new Date(event.StartTime).toLocaleString()}</p>
                          <p><strong>End Date:</strong> {new Date(event.EndDate).toLocaleString()}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}




              {/* pop up */}
              <div>
                {this.state.isPopupVisible && (
                  <Layer>
                    <Popup
                      className={styles.root}
                      role="dialog"
                      aria-modal="true"
                      onDismiss={this.hidePopup}
                    >
                      <Overlay
                        onClick={this.hidePopup}
                      />
                      <FocusTrapZone>
                        <div
                          role="document"
                          className={styles.content}
                        >
                          <div>
                            <div className={styles.cardgrid}>
                              <div className={styles.closeIcon} onClick={this.hidePopup}>
                                <Icon iconName="ChromeClose" />
                              </div>
                              <div className={styles.inputgroup}>
                                <TextField label="Title"
                                  value={this.state.EventTitle}
                                  onChange={this.onChangeEventTitle}
                                  className={styles.inputcontrol} />
                              </div>

                              <div className={styles.inputgroup}>
                                <Label>Start Date</Label>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gridColumnGap: '3px' }}>
                                  <DatePicker
                                    placeholder="Select a date..."
                                    value={this.state.StartDate}
                                    onSelectDate={this.onSelectDate}
                                    ariaLabel="Date picker"
                                  />
                                  <TimePicker
                                    placeholder="Select a time"
                                    value={this.state.StartTime}
                                    onChange={(ev, time) => this.onSelectTime(time!)} // Ensure time is not undefined
                                    ariaLabel="Time picker"
                                    useHour12={true}
                                  />
                                </div>
                              </div>

                              <div className={styles.inputgroup}>
                                <Label>End Date</Label>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gridColumnGap: '3px' }}>
                                  <DatePicker
                                    placeholder="Select a date..."
                                    value={this.state.EndDate}
                                    onSelectDate={this.onSelectEndDate}
                                    ariaLabel="Date picker"
                                  />
                                  <TimePicker
                                    placeholder="Select a time"
                                    value={this.state.EndTime}
                                    onChange={(ev, time) => this.onSelectEndTime(time!)}
                                    ariaLabel="Time picker"
                                    useHour12={true}

                                  />
                                </div>

                              </div>

                              <div className={styles.inputgroup}>
                                <TextField label="Description" multiline rows={3}
                                  value={this.state.EventDescription}
                                  onChange={this.onChangeDescription} />
                              </div>

                              <div className={styles.inputgroup}>
                                <Checkbox label="All Day"
                                  checked={this.state.AlldayNeeded}
                                  onChange={this.onAlldayNeededChange} />
                              </div>
                              <div className={styles.inputgroup}>
                                <TextField label="Where"
                                  value={this.state.Eventlocation}
                                  onChange={this.onChangeLocation}
                                  className={styles.inputcontrol} />
                              </div>
                              <div className={styles.inputgroup}>
                                <TextField label="Meeting Link"
                                  value={this.state.EventLink}
                                  onChange={this.onChangeLink}
                                  className={styles.inputcontrol} />
                              </div>

                              <div className={styles.inputgroup}>
                                <Dropdown
                                  label="Category"
                                  className={styles.inputcontrol}
                                  placeholder="Select"
                                  options={this.state.EventCategory}
                                  onChange={this.getOnchangeCategory}
                                  selectedKey={this.state.selectedCategory.key}
                                />
                              </div>



                              <div className={styles.inputgroup}>
                                <Label>Upload Image</Label>
                                <input
                                  type="file"
                                  accept="image/*"
                                  onChange={(e) => {
                                    const file = e.target.files ? e.target.files[0] : null;
                                    if (file) {
                                      console.log(file, " fle selected"); // You can handle file upload logic here
                                      this.setState({ selectedImage: file });
                                    }
                                  }}
                                />
                              </div>
                            </div>
                          </div>

                          <div className={styles.buttongroup}>
                            <PrimaryButton
                              onClick={this.onSubmitClick}
                              text="Submit"
                            // disabled={this.state.isOkButtonDisabled}

                            />
                            <DefaultButton onClick={this.hidePopup} >Cancel </DefaultButton>
                          </div>

                        </div>
                      </FocusTrapZone>
                    </Popup>
                  </Layer>
                )}
              </div>
            </div>
          </div >



        </div>
      </section>
    );
  }

}
