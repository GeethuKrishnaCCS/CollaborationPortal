import * as React from 'react';
import styles from './ProductAnniversary.module.scss';
import type { IProductAnniversaryProps, IProductAnniversaryState } from '../interfaces/IProductAnniversary';
import { ProductAnniversaryService } from '../services/ProductAnniversaryService';
// import { IIconProps, IconButton } from '@fluentui/react';
import * as moment from 'moment';

import { IIconProps } from '@fluentui/react';
import FilmstripView from '../views/FilmstripView';
import CompactView from '../views/CompactView';
import CardView from '../views/CardView';
import ListView from '../views/LIstView';

export default class ProductAnniversary extends React.Component<IProductAnniversaryProps, IProductAnniversaryState, {}> {
  private service: ProductAnniversaryService;

  constructor(props: IProductAnniversaryProps) {
    super(props);
    this.service = new ProductAnniversaryService(this.props.context, this.props.context.pageContext.web.serverRelativeUrl);
    this.state = {
      listItems: [],
      today: "",
      greetings: [],
      // RenderedGreetings: [],
      // Next: 5,
      // Count: 1
    };

    this.getAllUsersBirthdays = this.getAllUsersBirthdays.bind(this);
    this.getAnniversaryListItems = this.getAnniversaryListItems.bind(this);
    this.getEmployeeDetail = this.getEmployeeDetail.bind(this);

  }

  public async componentDidMount() {
    this.getAnniversaryListItems();
    // this.getAllUsersBirthdays();

    // const Listfields = await this.service.getListFields("Anniversary");
    // console.log('Listfields: ', Listfields);

  }


  public async getAllUsersBirthdays() {
    try {
      const graphClient = await this.props.context.msGraphClientFactory.getClient('3');

      // Step 1: Retrieve all users (only fetch the IDs)
      const usersResponse = await graphClient.api("users")
        .version("v1.0")
        .select("id")
        .get();

      const users = usersResponse.value;

      if (!users || users.length === 0) {
        console.log("No users found");
        return [];
      }

      // Step 2: For each user, fetch birthday and other details
      const birthdayPromises = users.map((user: any) =>
        graphClient.api(`users/${user.id}`)
          .version("v1.0")
          .select("birthday,displayName,jobTitle,mail")
          .get()
      );

      const usersWithBirthdays = await Promise.all(birthdayPromises);

      console.log('Users with Birthdays:', usersWithBirthdays);
      return usersWithBirthdays;

    } catch (error) {
      console.error('Error fetching user birthdays:', error);
      return [];
    }
  }


  public async getAnniversaryListItems() {
    try {
      // const url: string = `${this.props.context.pageContext.web.serverRelativeUrl}/Lists/Anniversary`;
      // const url: string = `${this.props.context.pageContext.web.serverRelativeUrl}/Lists/${this.props.Anniversary}`;
      const url: string = `${this.props.Anniversary}`;
      console.log('url: ', url);
      const listItem = await this.service.getSelectExpand(url, "*, Employee/ID, Employee/Title,Employee/EMail,Employee/JobTitle ", "Employee");

      const greetings: any[] = [];
      const currentDate = new Date();
      const todayDate = moment(currentDate).format('MMM DD');

      const greetingsPromises = listItem.map(async (item: any) => {
        try {
          const dateOfBirth = moment(item.Birthday).format('MMM DD');
          const dateOfJoining = moment(item.JoiningDate).format('MMM DD');
          const dateOfWedding = moment(item.WeddingDate).format('MMM DD');

          const employeeInfo = await this.service.getUser(item.Employee.ID);

          // Get the employee image URL here
          const employeeImage = this.getEmployeeDetail(item.Employee.Title, item.Employee.EMail).personImage;

          if (dateOfBirth === todayDate && this.props.BdayToggleValue) {
            greetings.push({ ...item, type: 'Birthday', employeeInfo, employeeImage });
          }
          if (dateOfJoining === todayDate && this.props.WorkToggleValue) {
            greetings.push({ ...item, type: 'Work Anniversary', employeeInfo, employeeImage });
          }
          if (dateOfWedding === todayDate && this.props.WeddingToggleValue) {
            greetings.push({ ...item, type: 'Wedding Anniversary', employeeInfo, employeeImage });
          }
        } catch (itemError) {
          console.error('Error in Fetching Data:', item, itemError);
        }
      });

      await Promise.all(greetingsPromises);
      this.setState({
        greetings: greetings,
        today: todayDate,
      });
      console.log('greetings: ', this.state.greetings);
    } catch (error) {
      console.error('Error fetching data:', error);
    }
  }

  public getEmployeeDetail(_name: string, _email: string) {
    // const personImageUrl = `${this.props.context.pageContext.web.absoluteUrl.replace(this.props.context.pageContext.web.serverRelativeUrl, '')}/_layouts/15/userphoto.aspx?size=L&accountname=${_email}`;

    return {
      displayName: _name,
      mail: _email,
      personImage: require('../assets/DefaultImage.png'), // Use fallback image
      // personImage: personImageUrl, // Use fallback image
    };
  }

  render() {
    const { hasTeamsContext, Cardlayout,  } = this.props;
    // const backicon: IIconProps = { iconName: 'ChevronLeftSmall' };
    // const nexticon: IIconProps = { iconName: 'ChevronRightSmall' };
    return (
      <section className={`${styles.productAnniversary} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* Heading */}
        <div className={styles.bdayheading} style={{ color: this.props.headingfontcolor, fontFamily: this.props.headingFontFamily }}>{this.props.description}</div>

        {/* Contents */}
        {Cardlayout === "Filmstrip" && (
          <FilmstripView
            greetings={this.state.greetings}
            backicon={{ iconName: 'ChevronLeftSmall' } as IIconProps}
            nexticon={{ iconName: 'ChevronRightSmall' } as IIconProps}
            contentfontcolor={this.props.contentfontcolor}
            headingfontcolor={this.props.headingfontcolor}
            headingFontFamily={this.props.headingFontFamily}
            contentFontfamily={this.props.contentFontfamily}
            bdayGreetingWish={this.props.bdayGreetingWish}
            WorkGreetingWish={this.props.WorkGreetingWish}
            weddingGreetingWish={this.props.weddingGreetingWish}
            context={this.props.context}

            ShapeForImages={this.props.ShapeForImages}
            displayListFields={this.props.displayListFields}
            AnniversaryListUrl = {this.props.Anniversary}
            
          />
        )}

        {Cardlayout === "Cards" && (
          <CardView
            greetings={this.state.greetings}
            backicon={{ iconName: 'ChevronLeftSmall' } as IIconProps}
            nexticon={{ iconName: 'ChevronRightSmall' } as IIconProps}
            contentfontcolor={this.props.contentfontcolor}
            headingfontcolor={this.props.headingfontcolor}
            bdayGreetingWish={this.props.bdayGreetingWish}
            WorkGreetingWish={this.props.WorkGreetingWish}
            weddingGreetingWish={this.props.weddingGreetingWish}
            context={this.props.context}

            ShapeForImages={this.props.ShapeForImages}
            displayListFields={this.props.displayListFields}
            AnniversaryListUrl = {this.props.Anniversary}

          />
        )}


        {Cardlayout === "List" && (
          <ListView
            greetings={this.state.greetings}
            headingfontcolor={this.props.headingfontcolor}
            contentfontcolor={this.props.contentfontcolor}
            bdayGreetingWish={this.props.bdayGreetingWish}
            WorkGreetingWish={this.props.WorkGreetingWish}
            weddingGreetingWish={this.props.weddingGreetingWish}
            AnniversayNoOfItemDisplay={this.props.AnniversayNoOfItemDisplay}
            context={this.props.context}
          />
        )}

        {Cardlayout === "Compact" && (
          <CompactView
            greetings={this.state.greetings}
            headingfontcolor={this.props.headingfontcolor}
            contentfontcolor={this.props.contentfontcolor}
            bdayGreetingWish={this.props.bdayGreetingWish}
            WorkGreetingWish={this.props.WorkGreetingWish}
            weddingGreetingWish={this.props.weddingGreetingWish}
            context={this.props.context}

          />
        )}

      </section>
    );
  }
}
