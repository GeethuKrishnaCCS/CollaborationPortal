import * as React from 'react';
import type { IProductAnniversaryModelProps, IProductAnniversaryProps } from '../interfaces';
import ProductAnniversary from "../components/ProductAnniversary";

export default class ProductAnniversaryModel extends React.Component<IProductAnniversaryModelProps> {
  public render(): React.ReactElement<IProductAnniversaryProps> {
    return (
      <ProductAnniversary
        description={this.props.description}
        isDarkTheme={this.props.isDarkTheme}
        environmentMessage={this.props.environmentMessage}
        hasTeamsContext={this.props.hasTeamsContext}
        context={this.props.context}
        userDisplayName={this.props.userDisplayName}
        Cardlayout={this.props.Cardlayout}
        headingfontcolor={this.props.headingfontcolor}
        contentfontcolor={this.props.contentfontcolor}
        StylesForCards={this.props.StylesForCards}
        dataSource={this.props.dataSource}
        Anniversary={this.props.Anniversary}
        AnniversayNoOfItemDisplay={this.props.AnniversayNoOfItemDisplay}
        BdayToggleValue={this.props.BdayToggleValue}
        WorkToggleValue={this.props.WorkToggleValue}
        WeddingToggleValue={this.props.WeddingToggleValue}
        bdayGreetingWish={this.props.bdayGreetingWish}
        WorkGreetingWish={this.props.WorkGreetingWish}
        weddingGreetingWish={this.props.weddingGreetingWish}
        SharePointLists={this.props.SharePointLists}
      
        ShapeForImages={this.props.ShapeForImages}
        collectionData={this.props.collectionData}
           
      >
      </ProductAnniversary>
    );
  }
}
