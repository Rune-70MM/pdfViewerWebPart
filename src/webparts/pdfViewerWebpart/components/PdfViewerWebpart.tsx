/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable @typescript-eslint/explicit-member-accessibility */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/jsx-no-bind */
//import styles from './PdfViewerWebpart.module.scss';

import * as React from 'react';
import { IPdfDataAccessState, IPdfViewerWebpartProps } from './IPdfViewerWebpartProps';
import { PdfViewerEntity } from '../../../sp/entities/pdf-viewer-webpart';
import { TextField, /*ChoiceGroup, Dropdown, IChoiceGroupOption, IDropdownOption, PrimaryButton*/ } from '@fluentui/react';
import { Stack, /*IStackStyles, IStackTokens*/ } from '@fluentui/react/lib/Stack';
import { IStackStyles, IStackTokens, Icon, PrimaryButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';


const stackStyles: IStackStyles = {
  root: {
    minWidth: '500px',
  },
};

const outerSpacingStackTokens: IStackTokens = {
  childrenGap: 30,
  padding: 10,

};

const textfieldSpacingStackTokens: IStackTokens = {
  childrenGap: 15,
  padding: 10,
};

export default class PdfViewerWebpart extends React.Component<IPdfViewerWebpartProps, IPdfDataAccessState> {

  public constructor(props: IPdfViewerWebpartProps) {
    super(props);

    this.state = {
      allItems: new Array<PdfViewerEntity>(),
      listitem: new PdfViewerEntity(),
      saving: false,
    }
  }

  async componentDidMount() {
    try {

      const listItems: PdfViewerEntity[] = await this.props.dataService.getAll();

      this.setState({
        allItems: listItems
      });

    } catch (error) {

    }
  }


  public render(): React.ReactElement<IPdfViewerWebpartProps> {

    const formIsNotValid = (): boolean => {

      if (!this.state.listitem.Title || !this.state.listitem.Category || !this.state.listitem.Status)
        return true;
      return false;
    }

    return (
      <>
        <Stack horizontal styles={stackStyles} tokens={outerSpacingStackTokens}>
          <span>
            <Stack styles={stackStyles} tokens={textfieldSpacingStackTokens}>
              <Stack tokens={{ childrenGap: 10 }}>
                <TextField label='Title' value={this.state.listitem.Title} onChange={(event, newValue) => {
                  this.setState({
                    ...this.state,
                    listitem: {
                      ...this.state.listitem,
                      Title: newValue
                    }
                  });
                }} />
                {/*
              <PeoplePicker context={this.props.context as any}
                titleText="User"
                personalSelectionLimit={1}
                showtooltip
                required
                defaultSelectedUsers={[this.state.listitem.Owner.Email ? this.state.listitem.Owner.Email : '']}
                onchange={(items: any[]) =>
                  {
                    let updatedListItem = { ...this.state.listitem};

                    if (items.length > 0 )
                    {
                      const Owner = new Person();
                      Owner.Email = items[0].secondaryText;
                      Owner.Title = items [0].text;

                      updatedListItem.Owner = Owner;
                     
                    }
                    else
                    {
                      updatedListItem.Owner = new Person();
                    }
                  this.state({ listItem: updatedListItem });

                  }}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
               
              <TextField label='Owner' value={this.state.listitem.Owner.Title} onChange={(event, newValue) => 
                  { 
                  this.setState({
                    ...this.state,
                    listitem:{
                        ...this.state.listitem,
                        Owner:[...this.state.listitem.Owner, ]
                    }
                  });
                }}/>
              */}
                <TextField label='Category' value={this.state.listitem.Category} onChange={(event, newValue) => {
                  this.setState({
                    ...this.state,
                    listitem: {
                      ...this.state.listitem,
                      Category: newValue
                    }
                  });
                }} />
                <TextField label='Status' value={this.state.listitem.Status} onChange={(event, newValue) => {
                  this.setState({
                    ...this.state,
                    listitem: {
                      ...this.state.listitem,
                      Status: newValue
                    }
                  });
                }} />
              </Stack>
            </Stack>
          </span>
          <span> {/*PDF - right side*/}
            <embed src='https://metafinanz.sharepoint.com/sites/GRP_DWPOnboardingTraining/JessicasSandbox/PDFViewer/Shared%20Documents/Examplepdf.pdf?#page=1&view=FitH&toolbar=0'//page=1&view=FitH&toolbar=0 
              type="application/pdf"
              width='500px'
              height="100%">
            </embed>
            {/*<object data="https://metafinanz.sharepoint.com/sites/GRP_DWPOnboardingTraining/JessicasSandbox/PDFViewer/Shared%20Documents/Examplepdf.pdf?#page=1&view=FitH&toolbar=0" //page=1&view=FitH&toolbar=0
                type="application/pdf" 
                width="100%"
                height="100%"       
        >
            <p>Alternative text - include a link <a href="https://metafinanz.sharepoint.com/sites/GRP_DWPOnboardingTraining/JessicasSandbox/PDFViewer/Shared%20Documents/Examplepdf.pdf">to the PDF!</a></p>
        </object> */}
          </span>
          <Stack horizontal horizontalAlign='end'>
            <PrimaryButton

              disabled={formIsNotValid() || this.state.saving}
              onClick={async () => {
                try {

                  this.setState({
                    saving: true
                  }, async () => {
                    //call service to add item

                    const listItemId = await this.props.dataService.addItem(this.state.listitem);

                    const allItems = await this.props.dataService.getAll();
                    listItemId;
                    this.setState({
                      saving: false,
                      allItems,

                    });

                    /* Nicht weiter benutzt
                    setTimeout(() => {
                      console.log('demo')
                    }, 1000);
                  */
                    this.setState({ saving: false });
                  })

                } catch (F) {
                  alert('Unable to save the data');
                }
              }}
            >
              <Stack horizontal tokens={{ childrenGap: 5 }}>
                {this.state.saving && <Spinner size={SpinnerSize.small} />}
                {!this.state.saving && <Icon iconName='Save' />}

                <span>Save</span>
              </Stack>
            </PrimaryButton>
          </Stack>

        </Stack>
        <Stack>
          <PrimaryButton

            disabled={formIsNotValid() || this.state.saving}
            onClick={async () => {
              try {

                this.setState({
                  saving: true
                }, async () => {
                  //call service to add item

                  const listItemId = await this.props.dataService.addItem(this.state.listitem);

                  const allItems = await this.props.dataService.getAll();
                  listItemId;
                  this.setState({
                    saving: false,
                    allItems,

                  });
                  this.setState({ saving: false });
                })

              } catch (F) {
                alert('Unable to save the data');
              }
            }}
          >
            <Stack horizontal tokens={{ childrenGap: 5 }}>
              {this.state.saving && <Spinner size={SpinnerSize.small} />}
              {!this.state.saving && <Icon iconName='Save' />}

              <span>Save</span>
            </Stack>
          </PrimaryButton>
        </Stack>
      </>)
  }
}

