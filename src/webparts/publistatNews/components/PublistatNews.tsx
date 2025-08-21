import * as React from 'react';
import styles from './PublistatNews.module.scss';
import { IPublistatNewsProps } from './IPublistatNewsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFieldInfo, Social, sp } from '@pnp/sp/presets/all';
import { DatePicker, DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogFooter, IColumn, Icon, IIconProps, PrimaryButton, SearchBox, Selection, SelectionMode, TextField } from 'office-ui-fabric-react';
import * as moment from 'moment';
import * as Excel from 'exceljs';
import { saveAs } from 'file-saver';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
import * as Chart from "../assets/js/Chart.min.js";

require('../assets/css/fabric.min.css');
require('../assets/css/style.css');

export interface IPublistatNewsState {
  AllNews: any;
  AddTagDialog: boolean;
  CurrentUserName: any;
  CurrentEmail: any;
  AddFormTag: any;
  AddFormSubscribed: boolean;
  AddFormSendNotifications: boolean;
  MyNewsTags: any;
  MySubscribedTags: any;
  MyNews: any;
  MyNewsFilterData: any;
  MySavedNews: any;
  FilterDialog: boolean;
  ExportData: any;
  FilteredExportData: any;
  searchText: string;
  startDate: any;
  endDate: any;
  selectedItems: any;
  selectionDetails: any;
  EmailDialog: boolean;
  RecevierEmailID: any;
  EmailSuccessDialog: boolean;
  SubscribedNewsCount: any;
  ShowMoreNews: boolean;
  hasMoreNews: boolean;
  position: number;
  totalCount: number;
  totalPages: number;
  currentPage: number;
  skipCount : number;
  totalItemCount : number;
  isLoading: boolean;
  pagedItems: any; 
}

let ctx;

const dialogContentProps = {
  title: 'Add News Tag',
};
const FilterdialogContentProps = {
  title: 'Export News',
};
const SendEmaildialogContentProps = {
  title: 'Send Mail',
};
const EmailSuccessDialogContentProps = {
  title: 'Mail Sent',
  subText: "The Mail has been sent successfully."
};


const follow: IIconProps = { iconName: 'Accept' };
const unfollow: IIconProps = { iconName: 'Cancel' };

const NotifyTrue: IIconProps = { iconName: 'RingerSolid' };
const NotifyFalse: IIconProps = { iconName: 'RingerOff' };

const Export: IIconProps = { iconName: 'DownloadDocument' };
const SendMail: IIconProps = { iconName: 'MailLowImportance' };

const FlowURL = {
  SendMail: "https://prod-161.westeurope.logic.azure.com:443/workflows/82e57b7dd0404f2a90cce36397e0ccb1/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=c2w2VR_Z54FydTdWmab5DkRNtudapFfFB1a5S_n9zTc",
};

const columns: IColumn[] = [
  {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 350,
    isResizable: true,
    onRender: (item) => {
      return <a href={item.Link} style={{ textDecoration: "none", color: "#006eb5" }}><div><span>{item.Source}: {item.Title}</span></div></a>;
    },
  },
  {
    key: 'Source',
    name: 'Source',
    fieldName: 'Source',
    minWidth: 50,
    maxWidth: 120,
    isResizable: true
  },
  {
    key: 'Pubdate',
    name: 'Publish Date',
    fieldName: 'Pubdate',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
    onRender: (item) => {
      return <span>{moment(new Date(item.Pubdate)).format('Do MMM')}</span>;
    },
  }

];

let XLcolums = [
  { header: "News Title", key: "Title" },
  { header: "Source", key: "Source" },
  { header: "Publish Date", key: "Pubdate" },
  { header: "URL", key: "Link" }
];


export default class PublistatNews extends React.Component<IPublistatNewsProps, IPublistatNewsState> {
  private _selection: Selection;
  constructor(props: IPublistatNewsProps, state: IPublistatNewsState) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
        });
      },
      getKey: this._getKey,
    });

    this.state = {
      AllNews: [],
      AddTagDialog: true,
      CurrentUserName: '',
      CurrentEmail: '',
      AddFormTag: '',
      AddFormSubscribed: true,
      AddFormSendNotifications: true,
      MyNewsTags: [],
      MySubscribedTags: [],
      MyNews: [],
      MyNewsFilterData: [],
      MySavedNews: [],
      FilterDialog: true,
      ExportData: [],
      FilteredExportData: [],
      searchText: '',
      startDate: '',
      endDate: '',
      selectedItems: [],
      selectionDetails: this._getSelectionDetails(),
      EmailDialog: true,
      RecevierEmailID: [],
      EmailSuccessDialog: true,
      SubscribedNewsCount: [],
      ShowMoreNews: true,
      hasMoreNews: true,
      position: 0,
      totalCount: 0,
      totalPages: 0,
      currentPage: 1,
      skipCount : 0,
      totalItemCount : 0,
      isLoading: false,
      pagedItems: null,
    };

  }

  public render(): React.ReactElement<IPublistatNewsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section id='PublistatNews'>
        <div className='ms-Grid'>
          <div className='ms-Grid-row'>
            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8'>
              <div className='d-flex-header'>
                <h2 className='Publistat-header'>Your News</h2>
                <PrimaryButton text='Export' iconProps={Export} onClick={() => this.setState({ FilterDialog: false })}></PrimaryButton>
              </div>
              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 ms-xl6'>
                <SearchBox placeholder="Search" className='new-search' onChange={(e) => this.SearchMyNews(e.target.value)} onClear={() => this.setState({ MyNews: this.state.ExportData })} />
              </div>
            </div>
          </div>
          <div className='ms-Grid-row flex-wrap-m'>
            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8'>
              <h4 className='Mynewstitle'>Your News</h4>
              {
                this.state.MyNews.length == 0 ?
                  <div style={{ textAlign: 'center', backgroundColor: '#ffffff', paddingBottom: '40px' }}>
                    <img className='NewsError-img' src={require('../assets/Images/newspaper.png')} />
                    <h4 className='NewsError-msg'>Welcome to your personalized news page.</h4>
                  </div>
                  : <></>
              }
              {
                this.state.MyNews.length > 0 && (
                  this.state.MyNews.map((item) => {
                    return (
                      <>
                        <div className='Publistat-Newcard'>
                          <p className='Publistat-Newsdate'>{moment(new Date(item.Pubdate)).format('Do MMM')} <span>- {item.Category}</span> </p>
                          <Icon className='SaveNews-Icon' onClick={() => this.MarkAsSave(item.Title, item.Link, item.Date, item.Source, item.ENTitle, item.ENDescription)} iconName='Pinned'></Icon>
                          <a className='NewsLink' href={item.Link} data-interception="off" target='_blank'>
                            <h4 className='Publistat-Newstitle'><span>{item.Source}</span>: {item.Title}</h4>
                            {item.ENTitle != '-' && item.ENTitle ? <h4 className='Publistat-NewsEntitle'>{item.ENTitle}</h4> : <></>}
                          </a>
                        </div>
                      </>
                    );
                  }
                  ))
              }

              {
                this.state.MyNews.length > 0 ? <>
                  {this.state.pagedItems && this.state.pagedItems.hasNext && (
                    // <button
                    //   onClick={() => this.LoadMoreNews()}
                    //   // disabled={this.state.isLoading}
                    //   className="px-4 py-2 bg-blue-600 text-white rounded-lg mt-4"
                    // >
                    //   {this.state.isLoading ? "Loading..." : "Load More"}
                    // </button>
                    <div style={{ textAlign: 'center', marginTop: '20px' }} >
                      <PrimaryButton style={{ cursor: 'pointer' }} text="Load More" onClick={() => this.LoadMoreNews()}></PrimaryButton>
                    </div>
                  )}
                </> : <></>
              }
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg4 ms-xl4'>
              <div className='Subscription-area Subscription-area-position'>
                <div className='ms-Grid-row'>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                    <h4 className='areatitle'>Subscribe News</h4>
                  </div>
                  <div className='ms-Grid-col ms-sm9 ms-md9 ms-lg9'>
                    <TextField className='AddTag-textfield' placeholder='Add new topics' onChange={(value) => this.setState({ AddFormTag: value.target["value"] })} value={this.state.AddFormTag} />
                  </div>
                  <div className='ms-Grid-col ms-sm3 ms-md3 ms-lg3'>
                    <div className='text-right mb-20'>
                      <PrimaryButton text='Add' onClick={() => this.AddTags()}></PrimaryButton>
                    </div>
                  </div>
                </div>

                {
                  this.state.MyNewsTags.length == 0 ?
                    <div style={{ textAlign: 'center' }}>
                      <img className='MyNewsTag-img' src={require('../assets/Images/bell.png')} />
                      <h4 className='MyNewsTag-msg'>Easily add and manage the topics that interest you.</h4>
                    </div>
                    : <></>
                }
                {
                  this.state.MyNewsTags.length > 0 && (
                    this.state.MyNewsTags.map((item) => {
                      return (
                        <>
                          <div className='ms-Grid-row Mytag-wrapper'>
                            <p className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 MyTag-title'>{item.NewsTag}</p>
                            <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4'>
                              <DefaultButton
                                toggle
                                checked={item.Subscribed == true ? true : false}
                                text={item.Subscribed == true ? 'Unfollow' : 'Follow'}
                                iconProps={item.Subscribed == true ? unfollow : follow}
                                onClick={() => this.UpdateSubscription(item.ID, item.Subscribed)}
                                allowDisabledFocus
                                className='SubscriptionBtn'
                              />
                            </div>
                            <div className='ms-Grid-col ms-sm2 ms-md2 ms-lg2 text-center'>
                              <DefaultButton
                                toggle
                                checked={item.SendNotifications == true ? true : false}
                                iconProps={item.SendNotifications == true ? NotifyTrue : NotifyFalse}
                                onClick={() => this.UpdateNotifications(item.ID, item.SendNotifications)}
                                allowDisabledFocus
                                title={item.SendNotifications == true ? "Don't Notify" : "Notify me"}
                                className='NotificationBtn'
                              />
                            </div>
                          </div>
                        </>
                      );
                    }
                    ))
                }


              </div>

              <div className='Subscription-area'>
                <div className='ms-Grid-row'>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                    <h4 className='areatitle'>Saved News</h4>
                  </div>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                    {
                      this.state.MySavedNews.length > 0 && (
                        this.state.MySavedNews.map((item) => {
                          return (
                            <>
                              <div className='Saved-Newcard'>
                                <p className='Saved-Newsdate'>{moment(new Date(item.Pubdate)).format('Do MMM')}</p>
                                <Icon iconName="Cancel" onClick={() => this.Unsave(item.ID)}></Icon>
                                <a className='NewsLink' href={item.Link} data-interception="off" target='_blank'>
                                  <h4 className='Saved-Newstitle'><span>{item.Source}</span>: {item.Title}</h4>
                                  {item.ENTitle != '-' && item.ENTitle ? <h4 className='Saved-NewsEntitle'>{item.ENTitle}</h4> : <></>}
                                </a>
                              </div>
                            </>
                          );
                        }
                        ))
                    }
                  </div>
                </div>
              </div>
              <div className='Subscription-area'>
                <div className='ms-Grid-row'>
                  <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                    <h4 className='areatitle'>Subscription Overview</h4>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 NewsTagGraph">
                    <div>
                      {
                        this.state.MySubscribedTags.length > 0 ?
                          <>
                            <canvas id="myChart" width="250" height="200"></canvas>
                          </> : <></>
                      }
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <Dialog
            hidden={this.state.AddTagDialog}
            onDismiss={() => this.setState({ AddTagDialog: true })}
            dialogContentProps={dialogContentProps}
            minWidth={450}
          >
            <div>
              <TextField label="Tag" onChange={(value) => this.setState({ AddFormTag: value.target["value"] })} value={this.state.AddFormTag} />
            </div>
            <DialogFooter>
              <PrimaryButton text="Add" onClick={() => this.AddTags()} />
              <DefaultButton onClick={() => this.setState({ AddTagDialog: true })} text="Cancel" />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={this.state.FilterDialog}
            onDismiss={() => this.setState({ FilterDialog: true, searchText: '', startDate: '', endDate: '', FilteredExportData: this.state.ExportData })}
            dialogContentProps={FilterdialogContentProps}
            minWidth={800}
          >
            <div>
              <div className='ms-Grid-row'>
                <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4'>
                  <SearchBox placeholder="Search" onChange={this.handleSearchChange} />
                </div>
                <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4'>
                  <DatePicker
                    placeholder="Select Start date..."
                    ariaLabel="Select a date"
                    onSelectDate={this.handleStartDateChange}
                    value={this.state.startDate}
                  />
                </div>
                <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4'>
                  <DatePicker
                    placeholder="Select End date..."
                    ariaLabel="Select a date"
                    onSelectDate={this.handleEndDateChange}
                    value={this.state.endDate}
                  />
                </div>
                <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 mt-10' style={{ textAlign: "right" }}>
                  <PrimaryButton text="Export" onClick={() => this.saveExcel()} iconProps={Export} />
                  <PrimaryButton text="Send Mail" className='ml-15' onClick={() => this.setState({ EmailDialog: false })} iconProps={SendMail} />
                </div>
              </div>
              <div>{this.state.selectionDetails}</div>
              <DetailsList
                items={this.state.FilteredExportData}
                // compact={isCompactMode}
                columns={columns}
                selectionMode={SelectionMode.multiple}
                setKey="multiple"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                // onItemInvoked={this._onItemInvoked}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="select row"
              />
            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EmailDialog}
            // onDismiss={toggleEmailDialog}
            dialogContentProps={SendEmaildialogContentProps}
            minWidth={400}
          >
            <TextField label="Recipient's Email" onChange={(e) => this.setState({ RecevierEmailID: e.target["value"] })} />
            <DialogFooter>
              <PrimaryButton onClick={() => this.triggerFlow(FlowURL.SendMail, this.state.selectedItems)} text="Send" />
              <DefaultButton onClick={() => this.setState({ EmailDialog: true })} text="Don't send" />
            </DialogFooter>
          </Dialog>
          <Dialog
            hidden={this.state.EmailSuccessDialog}
            // onDismiss={toggleHideDialog}
            dialogContentProps={EmailSuccessDialogContentProps}
            minWidth={400}
          >
            <DialogFooter>
              <PrimaryButton onClick={() => this.setState({ EmailSuccessDialog: true })} text="Ok" />
            </DialogFooter>
          </Dialog>

        </div>
      </section>
    );
  }

  public async componentDidMount(): Promise<void> {
    // await this.test();
    await this.HideNavigation();
    await this.GetCurrentUser();
    await this.GetMySubscribedTags();
    await this.GetMyTags();
    await this.GetSavedNews();
  }

  public async GetCurrentUser() {
    let user = await sp.web.currentUser.get();
    this.setState({ CurrentUserName: user.Title, CurrentEmail: user.Email });

  }

  public async GetMySubscribedTags() {
    sp.web.lists.getByTitle('User Prefrence').items.select('NewsTags').filter(`Email eq '${this.state.CurrentEmail}' and Subscribed eq 1`).get()
      .then((data) => {
        console.log(data);

        let tags = data.map(item => item.NewsTags).filter(tag => tag !== undefined && tag !== null);
        this.setState({ MySubscribedTags: tags });
        console.log(this.state.MySubscribedTags);

        this.GetNews();

      }
      )
      .catch((err) => {
        console.log(err);
      });
  }

  public async GetNews() {

    let items = [];
    let position = 0;
    // const pageSize = 2000;
    let AllData = [];
    try {
      // while (true) {
      //   const response = await sp.web.lists.getByTitle("News").items.select('Title', 'Link', 'Pubdate', 'Description', 'Date', 'Source', 'Newsgroup', 'Category', 'Newsguid','ENTitle','ENDescription', 'Sentiment','Reach','Topic','Spokesperson','Stakeholder','ArticleText','MediaType','Region','PublicationTime','PageNumber','Other').orderBy('Created',false).top(pageSize).skip(position).get();
      //   if (response.length === 0) {
      //     break;     
      //   }
      //   items = items.concat(response);
      //   position += pageSize;
      // }
      // console.log(`Total items retrieved: ${items.length}`);

      // let pagedItems = await sp.web.lists
      //   .getByTitle("News")
      //   .items
      //   .select('Title', 'Link', 'Pubdate', 'Description', 'Date', 'Source', 'Newsgroup', 'Category', 'Newsguid', 'ENTitle', 'ENDescription', 'Sentiment', 'Reach', 'Topic', 'Spokesperson', 'Stakeholder', 'ArticleText', 'MediaType', 'Region', 'PublicationTime', 'PageNumber', 'Other')
      //   .orderBy('Created', false)
      //   .top(pageSize)
      //   .getPaged();

      // items = items.concat(pagedItems.results);

      // while (pagedItems.hasNext) {
      //   pagedItems = await pagedItems.getNext();
      //   items = items.concat(pagedItems.results);
      // }

      //  let items = await sp.web.lists
      //   .getByTitle("News")
      //   .items
      //   .select('Title', 'Link', 'Pubdate', 'Description', 'Date', 'Source', 'Newsgroup', 'Category', 'Newsguid', 'ENTitle', 'ENDescription', 'Sentiment', 'Reach', 'Topic', 'Spokesperson', 'Stakeholder', 'ArticleText', 'MediaType', 'Region', 'PublicationTime', 'PageNumber', 'Other')
      //   .orderBy('Created', false)
      //   .top(pageSize)
      //   .get()

      // console.log(`Total items retrieved: ${items.length}`);


      const pageSize = 100;
      this.setState({ isLoading: true });

      let items = await sp.web.lists
        .getByTitle("News")
        .items
        .select('Title', 'Link', 'Pubdate', 'Description', 'Date', 'Source', 'Newsgroup', 'Category', 'Newsguid', 'ENTitle', 'ENDescription', 'Sentiment', 'Reach', 'Topic', 'Spokesperson', 'Stakeholder', 'ArticleText', 'MediaType', 'Region', 'PublicationTime', 'PageNumber', 'Other')
        .orderBy("Created", false)
        .top(pageSize)
        .getPaged();


      if (items.results.length > 0) {
        items.results.forEach((item, i) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Link: item.Link ? item.Link : "",
            Pubdate: item.Pubdate ? new Date(new Date(item.Date).setHours(new Date(item.Pubdate).getHours() + 2)).toISOString().split("T")[0] : "",
            Description: item.Description ? item.Description : "",
            Date: item.Date ? new Date(new Date(item.Date).setHours(new Date(item.Date).getHours() + 2)).toISOString().split("T")[0] : "",
            Source: item.Source ? item.Source : "",
            Newsgroup: item.Newsgroup ? item.Newsgroup : "",
            Category: item.Category ? item.Category : "",
            ENTitle: item.ENTitle ? item.ENTitle : "",
            ENDescription: item.ENDescription ? item.ENDescription : "",
            Sentiment: item.Sentiment ? item.Sentiment : "",
            Reach: item.Reach ? item.Reach : "",
            Topic: item.Topic ? item.Topic : "",
            Stakeholder: item.Stakeholder ? item.Stakeholder : "",
            Spokesperson: item.Spokesperson ? item.Spokesperson : "",
            ArticleText: item.ArticleText ? item.ArticleText : "",
            MediaType: item.MediaType ? item.MediaType : "",
            Region: item.Region ? item.Region : "",
            PublicationTime: item.PublicationTime ? item.PublicationTime : "",
            PageNumber: item.PageNumber ? item.PageNumber : "",
            Other: item.Other ? item.Other : "",
          });
        });
        this.setState({ AllNews: AllData });

        // let MySubTags = this.state.MySubscribedTags.map(tag => tag.toLowerCase());

        // let filteredData = this.state.AllNews.filter((x) => {
        //   let Title = x.Title;
        //   let Description = x.Description;
        //   let Category = x.Category;
        //   let Source = x.Source;
        //   let Newsgroup = x.Newsgroup;

        //   if (this.state.MySubscribedTags) {
        //     return MySubTags.some(tag => Title.includes(tag) || Description.includes(tag) || Category.includes(tag) || Source.includes(tag) || Newsgroup.includes(tag));
        //   }
        // });

        // console.log(filteredData);

        // ------------ * ---------

        //  let MySubTags = this.state.MySubscribedTags.map(tag => `\\b${tag.toLowerCase()}\\b`);

        //    let filteredData = this.state.AllNews.filter((x) => {
        //      let Title = x.Title.toLowerCase();
        //      let Description = x.Description.toLowerCase();
        //      let Category = x.Category.toLowerCase();
        //      let Source = x.Source.toLowerCase();
        //      let Newsgroup = x.Newsgroup.toLowerCase();
        //      let ENTitle = x.ENTitle.toLowerCase();
        //      let ENDescription = x.ENDescription.toLowerCase();

        //      if (this.state.MySubscribedTags) {
        //        return MySubTags.some(tag => 
        //          new RegExp(tag, 'i').test(Title) || 
        //          new RegExp(tag, 'i').test(Description) || 
        //          new RegExp(tag, 'i').test(Category) || 
        //          new RegExp(tag, 'i').test(Source) || 
        //          new RegExp(tag, 'i').test(Newsgroup) ||
        //          new RegExp(tag, 'i').test(ENTitle) ||
        //          new RegExp(tag, 'i').test(ENDescription)
        //        );
        //      }
        //    });
        //  this.setState({ MyNews: filteredData, MyNewsFilterData :filteredData });
        //  this.setState({ ExportData: filteredData,FilteredExportData: filteredData });

        // let MySubTags = this.state.MySubscribedTags.map(tag => tag.toLowerCase());

        // let filteredData = this.state.AllNews.filter((x) => {
        //     let Title = x.Title.toLowerCase();
        //     let Description = x.Description.toLowerCase();
        //     let Category = x.Category.toLowerCase();
        //     let Source = x.Source.toLowerCase();
        //     let Newsgroup = x.Newsgroup.toLowerCase();
        //     let ENTitle = x.ENTitle.toLowerCase();
        //     let ENDescription = x.ENDescription.toLowerCase();

        //     let allFields = [Title, Description, Category, Source, Newsgroup, ENTitle, ENDescription];

        //     // Helper function to check if a tag is found in any of the fields
        //     const isTagFound = (tag) => {
        //         let regex = new RegExp(`\\b${tag}\\b`, 'i');  // word boundary search
        //         return allFields.some(field => regex.test(field));
        //     };

        //     // Split subscribed tags into:
        //     // - Single tags (no AND condition)
        //     // - AND tags (contains " AND ")
        //     let singleTags = MySubTags.filter(tag => !tag.includes(' and '));
        //     let andTags = MySubTags.filter(tag => tag.includes(' and '));

        //     // OR Matching (for single tags)
        //     let orMatch = singleTags.some(tag => isTagFound(tag));

        //     // AND Matching (for "AND" combined tags)
        //     let andMatch = andTags.some(andTag => {
        //         let parts = andTag.split(' and ').map(tag => tag.trim());
        //         return parts.every(tag => isTagFound(tag));  // All parts must match
        //     });

        //     // Final result: show the item if either OR or AND condition matches
        //     return orMatch || andMatch;
        // });

        //   const filteredData = this.state.AllNews.filter((x) => {
        //     const fields = {
        //         title: x.Title.toLowerCase(),
        //         description: x.Description.toLowerCase(),
        //         category: x.Category.toLowerCase(),
        //         source: x.Source.toLowerCase(),
        //         newsgroup: x.Newsgroup.toLowerCase(),
        //         entitle: x.ENTitle.toLowerCase(),
        //         endescription: x.ENDescription.toLowerCase()
        //     };

        //     // Helper: match tag against all fields
        //     const evaluateTag = (rawTag: string): boolean => {
        //         rawTag = rawTag.trim().toLowerCase();

        //         // Wildcard support: e.g., binn*
        //         if (rawTag.includes('*')) {
        //             const regex = new RegExp(`\\b${rawTag.replace(/\*/g, '\\w*')}\\b`, 'i');
        //             return Object.values(fields).some(value => regex.test(value));  
        //         }

        //         // Basic word match (like "cs: politiek")
        //         const regex = new RegExp(`\\b${rawTag}\\b`, 'i');
        //         return Object.values(fields).some(value => regex.test(value));
        //     };

        //     // Logical expression evaluator (AND, OR, NOT, parentheses)
        //     const evaluateExpression = (expr: string): boolean => {
        //         expr = expr.replace(/\s+/g, ' ').trim().toLowerCase();

        //         const tokenize = (input: string): string[] =>
        //             input.replace(/([()])/g, ' $1 ').split(/\s+/).filter(Boolean);

        //         const toRPN = (tokens: string[]): string[] => {
        //             const output: string[] = [];
        //             const ops: string[] = [];
        //             const prec: any = { or: 1, and: 2, not: 3 };

        //             tokens.forEach(token => {
        //                 if (['and', 'or', 'not'].includes(token)) {
        //                     while (
        //                         ops.length &&
        //                         ops[ops.length - 1] !== '(' &&
        //                         prec[ops[ops.length - 1]] >= prec[token]
        //                     ) {
        //                         output.push(ops.pop()!);
        //                     }
        //                     ops.push(token);
        //                 } else if (token === '(') {
        //                     ops.push(token);
        //                 } else if (token === ')') {
        //                     while (ops.length && ops[ops.length - 1] !== '(') {
        //                         output.push(ops.pop()!);
        //                     }
        //                     ops.pop(); // Remove '('
        //                 } else {
        //                     output.push(token);
        //                 }
        //             });

        //             while (ops.length) output.push(ops.pop()!);
        //             return output;
        //         };

        //         const evalRPN = (rpn: string[]): boolean => {
        //             const stack: boolean[] = [];

        //             rpn.forEach(token => {
        //                 if (token === 'not') {
        //                     const val = stack.pop()!;
        //                     stack.push(!val);
        //                 } else if (token === 'and') {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a && b);
        //                 } else if (token === 'or') {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a || b);
        //                 } else {
        //                     stack.push(evaluateTag(token));
        //                 }
        //             });

        //             return stack[0];
        //         };

        //         return evalRPN(toRPN(tokenize(expr)));
        //     };

        //     // Final evaluation for each tag expression
        //     return this.state.MySubscribedTags.some(tagExpr => evaluateExpression(tagExpr));
        // });


        //   const filteredData = this.state.AllNews.filter((x) => {
        //     const fields = {
        //         title: x.Title,
        //         description: x.Description,
        //         category: x.Category,
        //         source: x.Source,
        //         newsgroup: x.Newsgroup,
        //         entitle: x.ENTitle,
        //         endescription: x.ENDescription
        //     };

        //     const searchableText = Object.values(fields).join(" ");
        //     const searchableTextLower = searchableText.toLowerCase();

        //     const matchKeyword = (rawWord: string): boolean => {
        //         rawWord = rawWord.trim();
        //         let caseSensitive = false;

        //         if (rawWord.toLowerCase().startsWith("cs:")) {
        //             caseSensitive = true;
        //             rawWord = rawWord.slice(3).trim();
        //         }

        //         const targetText = caseSensitive ? searchableText : searchableTextLower;
        //         const word = caseSensitive ? rawWord : rawWord.toLowerCase();

        //         // Wildcard support
        //         if (word.includes("*")) {
        //             const pattern = word.replace(/\*/g, "\\w*");
        //             const regex = new RegExp(`\\b${pattern}\\b`, caseSensitive ? "" : "i");
        //             return regex.test(targetText);
        //         }

        //         // Whole word match
        //         const regex = new RegExp(`\\b${word}\\b`, caseSensitive ? "" : "i");
        //         return regex.test(targetText);
        //     };

        //     const evaluateExpression = (expr: string): boolean => {
        //         const tokens = expr
        //             .replace(/([()])/g, " $1 ")
        //             .trim()
        //             .split(/\s+/)
        //             .filter(Boolean);

        //         const precedence: Record<string, number> = {
        //             or: 1,
        //             and: 2,
        //             not: 3
        //         };

        //         const toRPN = (tokens: string[]): string[] => {
        //             const output: string[] = [];
        //             const operators: string[] = [];

        //             tokens.forEach(token => {
        //                 const lower = token.toLowerCase();
        //                 if (["and", "or", "not"].includes(lower)) {
        //                     while (
        //                         operators.length &&
        //                         operators[operators.length - 1] !== "(" &&
        //                         precedence[operators[operators.length - 1].toLowerCase()] >= precedence[lower]
        //                     ) {
        //                         output.push(operators.pop()!);
        //                     }
        //                     operators.push(token);
        //                 } else if (token === "(") {
        //                     operators.push(token);
        //                 } else if (token === ")") {
        //                     while (operators.length && operators[operators.length - 1] !== "(") {
        //                         output.push(operators.pop()!);
        //                     }
        //                     operators.pop(); // remove "("
        //                 } else {
        //                     output.push(token);
        //                 }
        //             });

        //             while (operators.length) {
        //                 output.push(operators.pop()!);
        //             }

        //             return output;
        //         };

        //         const evalRPN = (rpn: string[]): boolean => {
        //             const stack: boolean[] = [];

        //             rpn.forEach(token => {
        //                 const lower = token.toLowerCase();
        //                 if (lower === "not") {
        //                     const val = stack.pop()!;
        //                     stack.push(!val);
        //                 } else if (lower === "and") {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a && b);
        //                 } else if (lower === "or") {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a || b);
        //                 } else {
        //                     stack.push(matchKeyword(token));
        //                 }
        //             });

        //             return stack[0];
        //         };

        //         const rpn = toRPN(tokens);
        //         return evalRPN(rpn);
        //     };

        //     // Now we loop through subscribed tags and apply NOT step-by-step
        //     return this.state.MySubscribedTags.some(tagExpr => {
        //         // Check if it includes NOT
        //         const parts = tagExpr.toLowerCase().split(/\s+not\s+/);

        //         if (parts.length === 2) {
        //             const includeExpr = parts[0].trim();
        //             const excludeExpr = parts[1].trim();

        //             const includeMatch = evaluateExpression(includeExpr);
        //             const excludeMatch = evaluateExpression(excludeExpr);

        //             return includeMatch && !excludeMatch;
        //         }

        //         // Normal evaluation (no NOT involved)
        //         return evaluateExpression(tagExpr);
        //     });
        // });

        //   const filteredData = this.state.AllNews.filter((x) => {
        //     const fields = {
        //         title: x.Title,
        //         description: x.Description,
        //         category: x.Category,
        //         source: x.Source,
        //         newsgroup: x.Newsgroup,
        //         entitle: x.ENTitle,
        //         endescription: x.ENDescription
        //     };

        //     const searchableText = Object.values(fields).join(" ");
        //     const searchableTextLower = searchableText.toLowerCase();

        //     const matchKeyword = (rawWord: string): boolean => {
        //         rawWord = rawWord.trim();
        //         let caseSensitive = false;

        //         // Check if the word starts with "CS:" and preserve case sensitivity
        //         if (rawWord.startsWith("CS:")) {
        //             caseSensitive = true;
        //             rawWord = rawWord.slice(3).trim(); // remove "CS:"
        //         }

        //         const targetText = caseSensitive ? searchableText : searchableTextLower;
        //         const word = caseSensitive ? rawWord : rawWord.toLowerCase();

        //         // Wildcard support
        //         if (word.includes("*")) {
        //             const pattern = word.replace(/\*/g, "\\w*");
        //             const regex = new RegExp(`\\b${pattern}\\b`, caseSensitive ? "" : "i");
        //             return regex.test(targetText);
        //         }

        //         // Whole word match
        //         const regex = new RegExp(`\\b${word}\\b`, caseSensitive ? "" : "i");
        //         return regex.test(targetText);
        //     };

        //     // Tokenizer that handles CS: Politiek as a single token
        //     const tokenizeExpression = (expr: string): string[] => {
        //         const regex = /(?:CS:\s*\w+)|\w+|\(|\)|AND|OR|NOT/gi;
        //         return [...expr.matchAll(regex)].map(match => match[0]);
        //     };

        //     const evaluateExpression = (expr: string): boolean => {
        //         const tokens = tokenizeExpression(expr); // Use custom tokenizer

        //         const precedence: Record<string, number> = {
        //             or: 1,
        //             and: 2,
        //             not: 3
        //         };

        //         const toRPN = (tokens: string[]): string[] => {
        //             const output: string[] = [];
        //             const operators: string[] = [];

        //             tokens.forEach(token => {
        //                 const lower = token.toLowerCase();
        //                 if (["and", "or", "not"].includes(lower)) {
        //                     while (
        //                         operators.length &&
        //                         operators[operators.length - 1] !== "(" &&
        //                         precedence[operators[operators.length - 1].toLowerCase()] >= precedence[lower]
        //                     ) {
        //                         output.push(operators.pop()!);
        //                     }
        //                     operators.push(token);
        //                 } else if (token === "(") {
        //                     operators.push(token);
        //                 } else if (token === ")") {
        //                     while (operators.length && operators[operators.length - 1] !== "(") {
        //                         output.push(operators.pop()!);
        //                     }
        //                     operators.pop(); // remove "("
        //                 } else {
        //                     output.push(token);
        //                 }
        //             });

        //             while (operators.length) {
        //                 output.push(operators.pop()!);
        //             }

        //             return output;
        //         };

        //         const evalRPN = (rpn: string[]): boolean => {
        //             const stack: boolean[] = [];

        //             rpn.forEach(token => {
        //                 const lower = token.toLowerCase();
        //                 if (lower === "not") {
        //                     const val = stack.pop()!;
        //                     stack.push(!val);
        //                 } else if (lower === "and") {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a && b);
        //                 } else if (lower === "or") {
        //                     const b = stack.pop()!;
        //                     const a = stack.pop()!;
        //                     stack.push(a || b);
        //                 } else {
        //                     stack.push(matchKeyword(token));
        //                 }
        //             });

        //             return stack[0];
        //         };

        //         const rpn = toRPN(tokens);
        //         return evalRPN(rpn);
        //     };

        //     // Now we loop through subscribed tags and apply NOT step-by-step
        //     return this.state.MySubscribedTags.some(tagExpr => {
        //         // Check if it includes NOT
        //         const parts = tagExpr.toLowerCase().split(/\s+not\s+/);

        //         if (parts.length === 2) {
        //             const includeExpr = parts[0].trim();
        //             const excludeExpr = parts[1].trim();

        //             const includeMatch = evaluateExpression(includeExpr);
        //             const excludeMatch = evaluateExpression(excludeExpr);

        //             return includeMatch && !excludeMatch;
        //         }

        //         // Normal evaluation (no NOT involved)
        //         return evaluateExpression(tagExpr);
        //     });
        // });

        const filteredData = this.state.AllNews.filter((x) => {
          const fields = {
            title: x.Title,
            description: x.Description,
            category: x.Category,
            source: x.Source,
            newsgroup: x.Newsgroup,
            entitle: x.ENTitle,
            endescription: x.ENDescription,
            Sentiment: x.Sentiment.toLowerCase(),
            Reach: x.Reach.toLowerCase(),
            Topic: x.Topic.toLowerCase(),
            Spokesperson: x.Spokesperson.toLowerCase(),
            Stakeholder: x.Stakeholder.toLowerCase(),
            ArticleText: x.ArticleText.toLowerCase(),
            MediaType: x.MediaType.toLowerCase(),
            Region: x.Region.toLowerCase(),
            PublictionTime: x.PublicationTime.toLowerCase(),
            PageNumber: x.PageNumber.toLowerCase(),
            Other: x.Other.toLowerCase(),
          };

          const searchableText = Object.values(fields).join(" ");
          const searchableTextLower = searchableText.toLowerCase();

          const matchKeyword = (rawWord: string): boolean => {
            rawWord = rawWord.trim();
            let caseSensitive = false;

            if (rawWord.startsWith("CS:")) {
              caseSensitive = true;
              rawWord = rawWord.slice(3).trim();
            }

            const targetText = caseSensitive ? searchableText : searchableTextLower;
            let word = caseSensitive ? rawWord : rawWord.toLowerCase();

            // Support wildcard at the end like 'binn*'
            let regex;
            if (word.endsWith("*")) {
              const prefix = word.slice(0, -1).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // escape special chars
              regex = new RegExp(`\\b${prefix}\\w*`, caseSensitive ? "" : "i");
            } else {
              const escapedWord = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
              regex = new RegExp(`\\b${escapedWord}\\b`, caseSensitive ? "" : "i");
            }

            return regex.test(targetText);
          };

          const shouldTokenize = (expr: string): boolean => {
            return /(AND|OR|NOT|\*|CS:)/i.test(expr);
          };

          const tokenizeExpression = (expr: string): string[] => {
            const regex = /CS:\s*\w+\*?|\w+\*?|\(|\)|AND|OR|NOT/gi;
            return [...expr.matchAll(regex)].map(match => match[0].trim());
          };


          const evaluateExpression = (expr: string): boolean => {
            if (!shouldTokenize(expr)) {
              // No special syntax detected, just treat as plain keyword match
              return matchKeyword(expr);
            }

            const tokens = tokenizeExpression(expr);
            const precedence: Record<string, number> = {
              or: 1,
              and: 2,
              not: 3
            };

            const toRPN = (tokens: string[]): string[] => {
              const output: string[] = [];
              const operators: string[] = [];

              tokens.forEach(token => {
                const lower = token.toLowerCase();
                if (["and", "or", "not"].includes(lower)) {
                  while (
                    operators.length &&
                    operators[operators.length - 1] !== "(" &&
                    precedence[operators[operators.length - 1].toLowerCase()] >= precedence[lower]
                  ) {
                    output.push(operators.pop()!);
                  }
                  operators.push(token);
                } else if (token === "(") {
                  operators.push(token);
                } else if (token === ")") {
                  while (operators.length && operators[operators.length - 1] !== "(") {
                    output.push(operators.pop()!);
                  }
                  operators.pop(); // remove "("
                } else {
                  output.push(token);
                }
              });

              while (operators.length) {
                output.push(operators.pop()!);
              }

              return output;
            };

            const evalRPN = (rpn: string[]): boolean => {
              const stack: boolean[] = [];

              rpn.forEach(token => {
                const lower = token.toLowerCase();
                if (lower === "not") {
                  const val = stack.pop()!;
                  stack.push(!val);
                } else if (lower === "and") {
                  const b = stack.pop()!;
                  const a = stack.pop()!;
                  stack.push(a && b);
                } else if (lower === "or") {
                  const b = stack.pop()!;
                  const a = stack.pop()!;
                  stack.push(a || b);
                } else {
                  stack.push(matchKeyword(token));
                }
              });

              return stack[0];
            };

            const rpn = toRPN(tokens);
            return evalRPN(rpn);
          };

          // Now we loop through subscribed tags and apply NOT step-by-step
          return this.state.MySubscribedTags.some(tagExpr => {
            // Check if it includes NOT
            const parts = tagExpr.toLowerCase().split(/\s+not\s+/);

            if (parts.length === 2) {
              const includeExpr = parts[0].trim();
              const excludeExpr = parts[1].trim();

              const includeMatch = evaluateExpression(includeExpr);
              const excludeMatch = evaluateExpression(excludeExpr);

              return includeMatch && !excludeMatch;
            }

            // Normal evaluation (no NOT involved)
            return evaluateExpression(tagExpr);
          });
        });
        this.setState({ MyNews: filteredData, MyNewsFilterData: filteredData, pagedItems: items, });
        this.setState({ ExportData: filteredData, FilteredExportData: filteredData });

        await this.GetNewsGraph();
      }
    } catch (error) {
      console.error(error);
    }


  }

  public async LoadMoreNews() {
    try {
      if (this.state.pagedItems && this.state.pagedItems.hasNext) {
        this.setState({ isLoading: true });

        const nextPage = await this.state.pagedItems.getNext();
        // const newData = this.formatItems(nextPage.results);

           const newData = nextPage.results.filter((x) => {
          const fields = {
            title: x.Title,
            description: x.Description,
            category: x.Category,
            source: x.Source,
            newsgroup: x.Newsgroup,
            entitle: x.ENTitle,
            endescription: x.ENDescription,
            Sentiment: x.Sentiment ? x.Sentiment.toLowerCase() : '',
            Reach: x.Reach ? x.Reach.toLowerCase() : '',
            Topic: x.Topic ? x.Topic.toLowerCase() : '',
            Spokesperson: x.Spokesperson ? x.Spokesperson.toLowerCase() : '',
            Stakeholder: x.Stakeholder ? x.Stakeholder.toLowerCase() : '',
            ArticleText: x.ArticleText ? x.ArticleText.toLowerCase() : '',
            MediaType: x.MediaType ? x.MediaType.toLowerCase() : '',
            Region: x.Region ? x.Region.toLowerCase() : '',
            PublictionTime: x.PublicationTime ? x.PublicationTime.toLowerCase() : '',
            PageNumber: x.PageNumber ? x.PageNumber.toLowerCase() : '',
            Other: x.Other ? x.Other.toLowerCase() :'',
          };

          const searchableText = Object.values(fields).join(" ");
          const searchableTextLower = searchableText.toLowerCase();

          const matchKeyword = (rawWord: string): boolean => {
            rawWord = rawWord.trim();
            let caseSensitive = false;

            if (rawWord.startsWith("CS:")) {
              caseSensitive = true;
              rawWord = rawWord.slice(3).trim();
            }

            const targetText = caseSensitive ? searchableText : searchableTextLower;
            let word = caseSensitive ? rawWord : rawWord.toLowerCase();

            // Support wildcard at the end like 'binn*'
            let regex;
            if (word.endsWith("*")) {
              const prefix = word.slice(0, -1).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // escape special chars
              regex = new RegExp(`\\b${prefix}\\w*`, caseSensitive ? "" : "i");
            } else {
              const escapedWord = word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
              regex = new RegExp(`\\b${escapedWord}\\b`, caseSensitive ? "" : "i");
            }

            return regex.test(targetText);
          };

          const shouldTokenize = (expr: string): boolean => {
            return /(AND|OR|NOT|\*|CS:)/i.test(expr);
          };

          const tokenizeExpression = (expr: string): string[] => {
            const regex = /CS:\s*\w+\*?|\w+\*?|\(|\)|AND|OR|NOT/gi;
            return [...expr.matchAll(regex)].map(match => match[0].trim());
          };


          const evaluateExpression = (expr: string): boolean => {
            if (!shouldTokenize(expr)) {
              // No special syntax detected, just treat as plain keyword match
              return matchKeyword(expr);
            }

            const tokens = tokenizeExpression(expr);
            const precedence: Record<string, number> = {
              or: 1,
              and: 2,
              not: 3
            };

            const toRPN = (tokens: string[]): string[] => {
              const output: string[] = [];
              const operators: string[] = [];

              tokens.forEach(token => {
                const lower = token.toLowerCase();
                if (["and", "or", "not"].includes(lower)) {
                  while (
                    operators.length &&
                    operators[operators.length - 1] !== "(" &&
                    precedence[operators[operators.length - 1].toLowerCase()] >= precedence[lower]
                  ) {
                    output.push(operators.pop()!);
                  }
                  operators.push(token);
                } else if (token === "(") {
                  operators.push(token);
                } else if (token === ")") {
                  while (operators.length && operators[operators.length - 1] !== "(") {
                    output.push(operators.pop()!);
                  }
                  operators.pop(); // remove "("
                } else {
                  output.push(token);
                }
              });

              while (operators.length) {
                output.push(operators.pop()!);
              }

              return output;
            };

            const evalRPN = (rpn: string[]): boolean => {
              const stack: boolean[] = [];

              rpn.forEach(token => {
                const lower = token.toLowerCase();
                if (lower === "not") {
                  const val = stack.pop()!;
                  stack.push(!val);
                } else if (lower === "and") {
                  const b = stack.pop()!;
                  const a = stack.pop()!;
                  stack.push(a && b);
                } else if (lower === "or") {
                  const b = stack.pop()!;
                  const a = stack.pop()!;
                  stack.push(a || b);
                } else {
                  stack.push(matchKeyword(token));
                }
              });

              return stack[0];
            };

            const rpn = toRPN(tokens);
            return evalRPN(rpn);
          };

          // Now we loop through subscribed tags and apply NOT step-by-step
          return this.state.MySubscribedTags.some(tagExpr => {
            // Check if it includes NOT
            const parts = tagExpr.toLowerCase().split(/\s+not\s+/);

            if (parts.length === 2) {
              const includeExpr = parts[0].trim();
              const excludeExpr = parts[1].trim();

              const includeMatch = evaluateExpression(includeExpr);
              const excludeMatch = evaluateExpression(excludeExpr);

              return includeMatch && !excludeMatch;
            }

            // Normal evaluation (no NOT involved)
            return evaluateExpression(tagExpr);
          });
        });

        this.setState(prevState => ({
          AllNews: [...prevState.AllNews, ...nextPage.results],
          MyNews: [...prevState.MyNews, ...newData],
          MyNewsFilterData: [...prevState.MyNewsFilterData, ...newData],
          ExportData: [...prevState.ExportData, ...newData],
          FilteredExportData: [...prevState.FilteredExportData, ...newData],
          pagedItems: nextPage,
          isLoading: false,
        }));

        console.log(nextPage.results);
        await this.GetNewsGraph();

      }
    } catch (error) {
      console.error(error);
      this.setState({ isLoading: false });
    }
  }

  public SearchMyNews(searchText: string): void {
    const { MyNewsFilterData } = this.state;

    const filteredItems = MyNewsFilterData.filter(item =>
      item.Title.toLowerCase().includes(searchText.toLowerCase()) ||
      item.Source.toLowerCase().includes(searchText.toLowerCase()) ||
      item.Category.toLowerCase().includes(searchText.toLowerCase()) ||
      item.Description.toLowerCase().includes(searchText.toLowerCase()) ||
      item.ENTitle.toLowerCase().includes(searchText.toLowerCase()) ||
      item.ENDescription.toLowerCase().includes(searchText.toLowerCase())
    );

    // Update the state with the filtered items
    this.setState({ MyNews: filteredItems });
  }

  public async UpdateSubscription(ItemID, subscription) {
    await sp.web.lists.getByTitle("User Prefrence").items.getById(ItemID).delete();
    //  const Tags = await sp.web.lists.getByTitle("User Prefrence").items.getById(ItemID).update({
    //     Subscribed: subscription == true ? false : true ,
    //   }).catch((err) => {
    //     console.log(err); 
    //   });
    await this.GetMyTags();
    await this.GetMySubscribedTags();
    // this.componentDidMount();
  }

  public async UpdateNotifications(ItemID, SendNotifications) {
    const Tags = await sp.web.lists.getByTitle("User Prefrence").items.getById(ItemID).update({
      SendNotifications: SendNotifications == true ? false : true,
    }).catch((err) => {
      console.log(err);
    });
    this.GetMyTags();
  }

  public async GetMyTags() {
    sp.web.lists.getByTitle('User Prefrence').items.select('Title', 'Email', 'NewsTags', 'Subscribed', 'SendNotifications', 'Id').filter(`Email eq '${this.state.CurrentEmail}'`).get()
      .then((data) => {
        let AllData = [];
        console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              NewsTag: item.NewsTags ? item.NewsTags : "",
              Subscribed: item.Subscribed ? item.Subscribed : "",
              SendNotifications: item.SendNotifications ? item.SendNotifications : ""
            });
          });
        }
        this.setState({ MyNewsTags: AllData });
        console.log(this.state.MyNewsTags);
      }
      )
      .catch((err) => {
        console.log(err);
      });
  }

  public async AddTags() {

    if (this.state.AddFormTag.length == 0) {
      alert("Please enter tag.");
    }
    else {
      const Tags = await sp.web.lists.getByTitle("User Prefrence").items.add({
        Title: this.state.CurrentUserName,
        Email: this.state.CurrentEmail,
        NewsTags: this.state.AddFormTag,
        Subscribed: this.state.AddFormSubscribed,
        SendNotifications: this.state.AddFormSendNotifications,
      }
      ).catch((err) => {
        console.log(err);

      });
      this.setState({ AddFormTag: '' });
      this.GetMySubscribedTags();
      this.GetMyTags();
    }
  }

  public async MarkAsSave(Title, URL, Date, Source, ENTitle, ENDescription) {
    await sp.web.lists.getByTitle("Saved News").items.add({
      Title: Title,
      Link: URL,
      Source: Source,
      Pubdate: Date,
      ENTitle: ENTitle,
      ENDescription: ENDescription,
    }
    ).catch((err) => {
      console.log(err);
    });
    this.GetSavedNews();
  }

  public async GetSavedNews() {
    sp.web.lists.getByTitle('Saved News').items.select('Title', 'Link', 'Pubdate', 'Author/Title', 'Id', 'Source', 'ENTitle', 'ENDescription').expand('Author').filter(`Author/Title eq '${this.state.CurrentUserName}'`).orderBy('Pubdate', false).get()
      .then((data) => {
        let AllData = [];
        console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              Title: item.Title ? item.Title : "",
              Link: item.Link ? item.Link : "",
              Pubdate: item.Pubdate ? new Date(new Date(item.Pubdate).setHours(new Date(item.Pubdate).getHours() + 2)).toISOString().split("T")[0] : "",
              Source: item.Source ? item.Source : "",
              ENTitle: item.ENTitle ? item.ENTitle : "",
              ENDescription: item.ENDescription ? item.ENDescription : ""
            });
          });
        }
        this.setState({ MySavedNews: AllData });
        console.log(this.state.MySavedNews);
      }
      )
      .catch((err) => {
        console.log(err);
      });
  }

  public async Unsave(ID) {
    await sp.web.lists.getByTitle("Saved News").items.getById(ID).delete();
    this.GetSavedNews();
  }

  public normalizeDate = (date: Date): Date => {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
    const searchText = newValue || '';
    this.setState({ searchText }, this.applyFilters);
  }

  handleStartDateChange = (date: Date | null) => {
    this.setState({ startDate: date }, this.applyFilters);
  }

  handleEndDateChange = (date: Date | null) => {
    this.setState({ endDate: date }, this.applyFilters);
  }

  applyFilters = () => {
    const { ExportData, searchText, startDate, endDate } = this.state;
    const FilteredExportData = ExportData.filter(item => {
      const Title = item.Title || '';
      const Source = item.Source || '';
      const Description = item.Description || '';
      const date = this.normalizeDate(new Date(item.Pubdate));

      const matchesSearch = !searchText || (Title.toLowerCase().includes(searchText.toLowerCase()) || Source.toLowerCase().includes(searchText.toLowerCase()) || Description.toLowerCase().includes(searchText.toLowerCase()));
      const matchesStartDate = !startDate || date >= this.normalizeDate(startDate);
      const matchesEndDate = !endDate || date <= this.normalizeDate(endDate);

      return matchesSearch && matchesStartDate && matchesEndDate;
    });
    this.setState({ FilteredExportData });
    console.log(this.state.FilteredExportData);
  }

  private _getSelectionDetails() {
    const selectionCount = this._selection.getSelectedCount();

    let selecteditems = this._selection.getSelection();
    console.log(selecteditems);

    this.setState({ selectedItems: selecteditems });
    // console.log(this.state.selectedItems);

  }

  private _getKey(item: any, index?: number): string {
    return item.Title;
  }


  private saveExcel = async () => {

    const web = sp.web;
    const siteTitle = await web.select("Title").get();

    const workbook = new Excel.Workbook();

    if (this.state.selectedItems.length > 0) {
      try {
        const fileName = moment().format("DD/MM/YYYY HH:MM") + ' Publistat Excel Overview ' + siteTitle.Title;
        const worksheet = workbook.addWorksheet();

        // add worksheet columns
        // each columns contains header and its mapping key from data
        worksheet.columns = XLcolums;

        // updated the font for first row.
        worksheet.getRow(1).font = { bold: true, color: { argb: '00000000' } };
        worksheet.getRow(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'ffffffff' } // Blue color#002e6d
        };

        // loop through all of the columns and set the alignment with width.
        // worksheet.columns.forEach(column => {
        //   column.width = 20;
        //   column.alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        // });

        worksheet.columns = [
          { width: 70 }, { width: 30 }, { width: 15 }, { width: 50 }
        ];
        worksheet.getColumn(1).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(2).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(3).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(4).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };

        const oddRowColor = 'FFFFFF'; // Lighter shade
        const evenRowColor = 'fbfbfb'; // Darker shade
        const borderColor = 'aaaaaa'; // Dark border color FFBFDEF7

        // Loop through data and add each one to worksheet
        this.state.selectedItems.forEach((singleData: any, index: number) => {
          const row = worksheet.addRow(singleData);

          // Set fill color based on odd or even row
          const fillColor = index % 2 === 0 ? { argb: oddRowColor } : { argb: evenRowColor };
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: fillColor
          };

          // Set border color for each cell in the row
          row.eachCell((cell, colNumber) => {
            cell.border = {
              top: { style: 'thin', color: { argb: borderColor } },
              left: { style: 'thin', color: { argb: borderColor } },
              bottom: { style: 'thin', color: { argb: borderColor } },
              right: { style: 'thin', color: { argb: borderColor } },
            };
          });
        });

        // write the content using writeBuffer
        const buf = await workbook.xlsx.writeBuffer();

        // download the processed file
        saveAs(new Blob([buf]), `${fileName}.xlsx`);
      } catch (error) {
        console.error('Something Went Wrong', error.message);
      }
    } else {
      alert("Please select News you want to export, then click the 'Export' button.");
    }
  }

  public triggerFlow = (postURL, data) => {

    if (this.state.RecevierEmailID.length > 0) {
      if (this.state.selectedItems.length > 0) {
        this.setState({ EmailDialog: true, RecevierEmailID: "" });

        const mail = this.state.RecevierEmailID;
        const data1 = JSON.stringify({ data, mail });
        const body: string = data1;

        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');

        const httpClientOptions: IHttpClientOptions = {
          body: body,
          headers: requestHeaders
        };

        return this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions).then((response) => {
          console.log("Flow Triggered Successfully...");
          this.setState({ EmailDialog: true, EmailSuccessDialog: false });

        }).catch(error => {
          console.log(error);
        });

      }
      else {
        alert("Please select News you want to export, then click the 'Send Mail' button.");
      }
    }
    else {
      alert("Please Add the recipient's email address.");
    }

  }

  public async HideNavigation() {

    try {
      // Get current user's groups
      const userGroups = await sp.web.currentUser.groups();

      // Check if the user is in the Owners or Admins group
      const isAdmin = userGroups.some(group =>
        group.Title.indexOf("Owners") !== -1 ||
        group.Title.indexOf("Admins") !== -1
      );

      if (!isAdmin) {
        // Hide the navigation bar for non-admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: none;");
        }
      } else {
        // Show the navigation bar for admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: block;");
        }
      }
    } catch (error) {
      console.error("Error checking user permissions: ", error);
    }

  }

  public async GetNewsGraph() {

    const getLast7Days = (): string[] => {
      const dates: string[] = [];
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

      for (let i = 0; i < 7; i++) {
        const date = new Date();
        date.setDate(date.getDate() - i);

        const day = date.getDate();
        const month = months[date.getMonth()];

        const formattedDate = `${day} ${month}`;
        dates.push(formattedDate);
      }
      return dates;
    };

    // console.log(getLast7Days());

    console.log(getLast7Days());

    interface NewsItem {
      Pubdate: string; // Assuming Timestamp is a valid date string
    }

    let today: Date = new Date();
    today.setHours(0, 0, 0, 0); // Normalize to midnight

    let last7DaysCounts: number[] = [0, 0, 0, 0, 0, 0, 0];

    this.state.MyNews.forEach((newsItem: NewsItem) => {
      let newsDate: Date = new Date(newsItem.Pubdate);
      newsDate.setHours(0, 0, 0, 0); // Normalize time for accurate comparison

      let diffInDays: number = Math.floor((today.getTime() - newsDate.getTime()) / (1000 * 60 * 60 * 24)); // Difference in days

      if (diffInDays >= 0 && diffInDays < 7) {
        last7DaysCounts[6 - diffInDays] += 1; // Store in correct index (latest news at the end)
      }
    });

    // Debugging logs
    console.log("News Counts for Last 7 Days:", last7DaysCounts);
    console.log("Today's Date:", today.toDateString());
    console.log("AllNews Count:", this.state.AllNews.length);


    // let MySubTags = this.state.MySubscribedTags.map(
    //   (tag) => new RegExp(`\\b${tag.toLowerCase()}\\b`, "i")
    // );

    // let tagNewsCount = this.state.MySubscribedTags.map(() => 0); // Array to store counts

    // this.state.AllNews.forEach((x) => {
    //   let Title = x.Title.toLowerCase();
    //   let Description = x.Description.toLowerCase();
    //   let Category = x.Category.toLowerCase();
    //   let Source = x.Source.toLowerCase();
    //   let Newsgroup = x.Newsgroup.toLowerCase();

    //   MySubTags.forEach((regex, index) => {
    //     if (
    //       regex.test(Title) ||
    //       regex.test(Description) ||
    //       regex.test(Category) ||
    //       regex.test(Source) ||
    //       regex.test(Newsgroup)
    //     ) {
    //       tagNewsCount[index] += 1;
    //     }
    //   });
    // });

    // console.log("News Counts Array:", tagNewsCount);
    this.setState({ SubscribedNewsCount: last7DaysCounts });
    var yValues = this.state.SubscribedNewsCount;
    var xValues = getLast7Days().reverse();
    var barColors = "#006eb5";

    if (ctx) {
      ctx.destroy(); // Destroy the previous chart
    }

    ctx = new Chart("myChart", {
      type: "bar",
      data: {
        labels: xValues,
        datasets: [{
          backgroundColor: barColors,
          data: yValues,
        }]
      },
      options: {
        legend: { display: false },
        responsive: true,
        tickWidth: '10',
        title: {
          display: true,
          text: "Subscribed News"
        },
        scales: {
          yAxes: [{
            ticks: {
              beginAtZero: true,
              steps: 10
            }
          }],
          xAxes: [{
            ticks: {
              fontSize: 11
            }
          }]
        }
      }
    });

  }

}    
