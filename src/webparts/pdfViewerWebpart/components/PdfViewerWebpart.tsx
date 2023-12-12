/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable @typescript-eslint/explicit-member-accessibility */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/jsx-no-bind */
//import styles from './PdfViewerWebpart.module.scss';

//import { escape } from '@microsoft/sp-lodash-subset';

//import { PrincipalType } from '@pnp/sp';
//import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import * as React from 'react';
import { IPdfDataAccessState, IPdfViewerWebpartProps } from './IPdfViewerWebpartProps';
import { PdfViewerEntity } from '../../../sp/entities/pdf-viewer-webpart';
import { TextField, /*ChoiceGroup, Dropdown, IChoiceGroupOption, IDropdownOption, PrimaryButton*/ } from '@fluentui/react';
import { Stack, /*IStackStyles, IStackTokens*/} from '@fluentui/react/lib/Stack';
import { IStackStyles, IStackTokens, Icon, PrimaryButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';




const stackStyles: IStackStyles = {
  root: {
    //background: DefaultPalette.themeTertiary,
    //width: '500px',
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
/*
const choiceOptions: IChoiceGroupOption[] = [
  { key: 'A', text: 'Show TextField Project' },
  { key: 'B', text: 'Disable TextField Project' },
  { key: 'C', text: 'Show TextField Description'},
  { key: 'D', text: 'Hide TextField Description' },
];

const dropDownOptions: IDropdownOption[] = [
  { key: 'approved', text: 'Approved' },
  { key: 'rejected', text: 'Rejected' },
];

*/

export default class PdfViewerWebpart extends React.Component<IPdfViewerWebpartProps, IPdfDataAccessState> {
  
 

  public constructor(props: IPdfViewerWebpartProps)
  {
    super(props);

    this.state = {
      allItems: new Array<PdfViewerEntity>(),
      listitem: new PdfViewerEntity(),
      saving: false,
    }
  }
  
  
  async componentDidMount() 
    {
     try {
      
      const listItems: PdfViewerEntity[] = await this.props.dataService.getAll();

      this.setState({
        allItems: listItems
      });

     } catch (error) {
      
     } 
    }


  public render(): React.ReactElement<IPdfViewerWebpartProps> {
    /*const {
      userDisplayName,
      /*description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
    } = this.props;
    */
      const formIsNotValid = (): boolean =>
      {

        if(!this.state.listitem.Title || !this.state.listitem.Category || !this.state.listitem.Status)
        return true;

/*      So  schaut es einzeln geschrieben aus, man kann es auch wie oben zusammenfassen
        if(!this.state.listitem.Owner)
        return true;

        if(!this.state.listitem.InvoiceCategory)
        return true;

        if(!this.state.listitem.Status)
        return true;
*/
        return false;
      }

    return (
      <>
      <Stack horizontal styles={stackStyles} tokens={outerSpacingStackTokens}>
        <span>
          <Stack styles={stackStyles} tokens={textfieldSpacingStackTokens}>
            {/*<section className={`${styles.pdfViewerWebpart} ${hasTeamsContext ? styles.teams : ''}`}> 
              
              <div className={styles.welcome}>
                <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
                <h2>Well done, {escape(userDisplayName)}!</h2>
                <div>{environmentMessage}</div>
                <div>Web part property value: <strong>{escape(description)}</strong></div>
              </div>
              <div>*/}
            <Stack tokens={{ childrenGap: 10 }}>
              <TextField label='Title' value={this.state.listitem.Title} onChange={(event, newValue) => 
                  { 
                  this.setState({
                    ...this.state,
                    listitem:{
                        ...this.state.listitem,
                        Title: newValue
                    }
                  });
                }}/>
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
                <TextField label='Category' value={this.state.listitem.Category} onChange={(event, newValue) => 
                  { 
                  this.setState({
                    ...this.state,
                    listitem:{
                        ...this.state.listitem,
                        Category: newValue
                    }
                  });
                }}/>
                <TextField label='Status' value={this.state.listitem.Status} onChange={(event, newValue) => 
                            { 
                            this.setState({
                              ...this.state,
                              listitem:{
                                  ...this.state.listitem,
                                  Status: newValue
                              }
                            });
                          }}/>

                          
                        {/*
                        {this.state.allItems.length > 0 && <>  
                                  <table>
                          <thead>
                            <tr>
                              <th>Owner</th>   
                              <th>Category</th>
                              <th>Status</th>
                              </tr>
                          </thead>
                          <tbody>
                            {this.state.allItems.map((item: PdfViewerEntity ) =>
                            {
                              return <tr key={item.Id}>
                                <td>{item.Owner}</td>
                                <td>{item.Category}</td>
                                <td>{item.Status}</td>
                              </tr>
                            })}
                          </tbody>
                        </table>  */}
            </Stack>
          </Stack>
          </span>
            {/*
            <span>{this.state.listitem.Title}</span>
            </Stack>
            <h1>Hello from PdfViewerWebPart</h1>
    
            <Stack horizontal styles={stackStyles} tokens={outerSpacingStackTokens}>

          <span> {/*text fields - left side*
            <Stack styles={stackStyles} tokens={textfieldSpacingStackTokens}>
            <TextField>Test1</TextField>
            <TextField>Test2</TextField>
            <TextField>Test3</TextField>
            
            

            <span> ---- Dynamic Fields ---- </span> 


            <ChoiceGroup defaultSelectedKey="A" options={choiceOptions} label="Pick one" required={true} />

            <TextField label="Project"></TextField>
            (<TextField hidden = {true} label="Description"></TextField>)

            
            <Dropdown
                placeholder="Select an option"
                label="Approval Status"
                options={dropDownOptions}
            />
            
            <PrimaryButton>Update</PrimaryButton>
            </Stack>
        
            </span>
            */}
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
            onClick={async () =>
            {
                try
                {

                  this.setState({
                    saving:true
                  }, async ()=>{
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

                }catch (F)
                  {
                      alert('Unable to save the data');
                  }
                    }}
                >
              <Stack horizontal tokens={{ childrenGap: 5}}>
                {this.state.saving && <Spinner size={SpinnerSize.small}/>}
                {!this.state.saving && <Icon iconName='Save'/>}
                
                <span>Save</span>
              </Stack>
            </PrimaryButton>
        </Stack>

   </Stack> 
   <Stack>
            <PrimaryButton
            
            disabled={formIsNotValid() || this.state.saving}
            onClick={async () =>
            {
                try
                {

                  this.setState({
                    saving:true
                  }, async ()=>{
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

                }catch (F)
                  {
                      alert('Unable to save the data');
                  }
                    }}
                >
              <Stack horizontal tokens={{ childrenGap: 5}}>
                {this.state.saving && <Spinner size={SpinnerSize.small}/>}
                {!this.state.saving && <Icon iconName='Save'/>}
                
                <span>Save</span>
              </Stack>
            </PrimaryButton>
        </Stack>
  </>)
  }
}

