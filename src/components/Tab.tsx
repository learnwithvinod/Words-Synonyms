// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component<any, any> {
  constructor(props: any){
    super(props)

    this.state = {
      context: {}
    }
    this.onInputchange = this.onInputchange.bind(this);
  }

  onInputchange(event: { target: { name: any; value: any; }; }) {
    let suggestion = this.xmlhttpPost("http://localhost:8983/solr/happy-searching/select", event.target.value, event)
    this.setState({
      [event.target.name]: suggestion
    }); 
  }

  xmlhttpPost(this: any, strURL: string, value: string, event: { target: { name: any; value: any; }; } ) {
    let xmlHttpReq: XMLHttpRequest;
    let result: string;
    let self = this;
    xmlHttpReq = new XMLHttpRequest();
    xmlHttpReq.open('POST', strURL, true);
    xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xmlHttpReq.onreadystatechange = function() {
      if (xmlHttpReq.readyState == 4) {
        let resp: string = xmlHttpReq.responseText;
        let rsp = eval("("+resp+")")
        if(rsp.response && rsp.response.docs && rsp.response.docs.length > 0){
          result = rsp.response.docs[0].suggestion;
        }        
      }
      if(result){
        self.setState({
          [event.target.name]: result[0]
        });
      }
      else {
        self.setState({
          [event.target.name]: value
        });
      }
  }
  
    var params = this.getstandardargs().concat(this.getquerystring(value));
    var strData = params.join('&');
    xmlHttpReq.send(strData);
  }
  
  getstandardargs() {
    var params = [
        'wt=json'
        , 'indent=on'
        , 'hl=true'
        , 'hl.fl=name,features'
        ];
  
    return params;
  }
  getquerystring(value: string) {
    let qstr = 'q=word:' + value;
    return qstr;
  }
  
  // this function does all the work of parsing the solr response and updating the page.
  updatepage(str: string){
    var rsp = eval("("+str+")"); // use eval to parse Solr's JSON response
    var html= "test";
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount(){
    // Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      this.setState({
        context: context
      });
    });
    // Next steps: Error handling using the error object
  }

  render() {

      const userName = Object.keys(this.state.context).length > 0 ? this.state.context['upn'] : "";

      return (
      <div>
        <h3>Digital Spartans!</h3>
        <h3>Inclusion & Diversity</h3>
        <textarea name="word" onChange={this.onInputchange}></textarea>
        <div className="result">{this.state.word}</div>
      </div>

      );
  }
}
export default Tab;