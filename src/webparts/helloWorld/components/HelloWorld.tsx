import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
// import { IHelloWorldStates } from './IHelloWorldStates';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
export interface IHelloWorldStates {
 personName: string;  
}
 export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldStates> {
   public localPersonName=this.props.personName;
constructor(props:IHelloWorldProps){
  super(props);
  this.state={
    personName:this.props.personName
  };
}
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    
    return (

      <><div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.description}><b>SPFx Crud Operations</b></p>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item ID:</div>
                <input className={styles.fieldLabelinput} type="text" id='itemId'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input className={styles.fieldLabelinput} type="text" id='fullName'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Isactive</div>
                <select id='Isactive' className={styles.fieldLabelinput}>
                  <option value="true">Yes</option>
                  <option value="false">No</option>
                </select>
              </div>
              {/* <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Isactive</div>
                <input type="radio" id="Isactive" name="Isactive" value="true"></input>
                <label >Yes</label>
                <input type="radio" id="Isactive" name="Isactive" value="false"></input>
                <label >No</label>
              </div> */}

            

              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Joining Date</div>
                <input type="date" id='JoiningDate' className={styles.fieldLabelinput}></input>
              </div>


              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>WorkingShift</div>
                <select id='WorkingShift' className={styles.fieldLabelinput}>
                  <option>DayShift</option>
                  <option>NightShift</option>
                </select>
              </div>
              {/* <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Documents</div>
                <input type="file" id='Documents'></input>
              </div> */}

              <div className={styles.itemField}>
                <div className={styles.fieldLabel} >Comments</div>
                <input type="text" id='Comments' className={styles.fieldLabelinput}></input>
              </div>


              
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems"></div>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>Create</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>Read</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>Read All</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>Update</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>Delete</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      {/* <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.welcome}>
            <h2>Well done, {escape(userDisplayName)}!</h2>
            <div>{environmentMessage}</div>
            <div>Web part property value: <strong>{escape(description)}</strong></div>
          </div>
          <div>
            <p className={styles.helloWorld}>{escape(this.props.description)}</p>

            <span className={styles.helloWorld}><b>Happy Birthday</b></span>
            <span><i>{escape(this.localPersonName)}</i></span>

            <span className={styles.helloWorld}><b>Happy Birthday</b></span>
            <span><i>{escape(this.state.personName)}</i></span>

            <p className={styles.helloWorld}>{escape(this.props.listName)}</p>
            <a href="https://aka.ms/spfx">
              <span>Learn more</span>
            </a>

            <button onClick={() => this.setProps()}>setProps</button>
            <button onClick={() => this.setStateValue()}>setState</button>

          </div>
        
      
      </section> */}
        </>
    );
    
  }
  // Create Item
  
  private createItem = (): void => {
    const body: string = JSON.stringify({
      'Name': document.getElementById("fullName")['value'],
      'Isactive': document.getElementById("Isactive")['value'],
      'JoiningDate': document.getElementById("JoiningDate")['value'],
      'WorkingShift': document.getElementById("WorkingShift")['value'],
      // 'Documents': document.getElementById("Documents")['value'],
      'Comments': document.getElementById("Comments")['value']
    });
    
    
    this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SPFXCrud')/items`,
      SPHttpClient.configurations.v1, {
        
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
        
      },
      body: body
    })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Item created successfully with ID: ${responseJSON.ID}`);
          });
        } else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! Check the error in the browser console.`);
          });
        }
      }).catch(error => {
        console.log(error);
      });
  }
 
  
// Get Item by ID
  private getItemById = (): void => {
    const id: number = document.getElementById('itemId')['value'];
    if (id > 0) {
      this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('SPFXCrud')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              document.getElementById('fullName')['value'] = responseJSON.Name;
              document.getElementById('Isactive')['value'] = responseJSON.Isactive;
              document.getElementById('JoiningDate')['value'] = responseJSON.JoiningDate;
              document.getElementById('WorkingShift')['value'] = responseJSON.WorkingShift;
              // document.getElementById('Documents')['value'] = responseJSON.Documents;
              document.getElementById('Comments')['value'] = responseJSON.Comments;
            });
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Something went wrong! Check the error in the browser console.`);
            });
          }
        }).catch(error => {
          console.log(error);
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
 
  
// Get all items
  private getAllItems = (): void => {
    this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('SPFXCrud')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            var html = `<table><tr style="color: #F22BB0;"><th>ID</th><th style="width: 87px; display: flex;">Full Name</th><th>Isactive</th><th style="width: 87px; display: flex;">Joining Date</th><th>WorkingShift</th><th>Comments</th></tr>`;
            responseJSON.value.map((item, index) => {
              html += `<tr><td>${item.ID}</td><td>${item.Name}</td><td>${item.Isactive}</td>
              <td>${item.JoiningDate.split("T")[0]}</td><td>${item.WorkingShift}</td><td>${item.Comments}</td></li>`;
            });
            html += `</table>`;
            document.getElementById("allItems").innerHTML = html;
            console.log(responseJSON);
          });
        } else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! Check the error in the browser console.`);
          });
        }
      }).catch(error => {
        console.log(error);
      });
  }
 
  
// Update Item
  private updateItem = (): void => {
    const id: number = document.getElementById('itemId')['value'];
    const body: string = JSON.stringify({
      'Name': document.getElementById("fullName")['value'],
      'Isactive': document.getElementById("Isactive")['value'],
      'JoiningDate': document.getElementById("JoiningDate")['value'],
      'WorkingShift': document.getElementById("WorkingShift")['value'],
      // 'Documents': document.getElementById("Documents")['value'],
      'Comments': document.getElementById("Comments")['value']
    });
    if (id > 0) {
      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SPFXCrud')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item with ID: ${id} updated successfully!`);
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Something went wrong! Check the error in the browser console.`);
            });
          }
        }).catch(error => {
          console.log(error);
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
 
  
// Delete Item
  private deleteItem = (): void => {
    const id: number = parseInt(document.getElementById('itemId')['value']);
    if (id > 0) {
      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SPFXCrud')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item ID: ${id} deleted successfully!`);
          }
          else {
            alert(`Something went wrong!`);
            console.log(response.json());
          }
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
    // public setProps(){
    //   console.log("Before: ",this.localPersonName)
    //   this.localPersonName = "Manish";
    //   console.log("After: ",this.localPersonName)
    // }
    // public setStateValue(){
    //   console.log("Before: ",this.state.personName)
    //   this.setState({
    //     personName:"Manish"
    //   },()=>{console.log("After: ",this.state.personName)});
      
    // }
    
  }

