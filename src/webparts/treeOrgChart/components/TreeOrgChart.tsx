
import * as React from "react";
import {HoverCard ,IExpandingCardProps} from 'office-ui-fabric-react';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import styles from "./TreeOrgChart.module.scss";
import { ITreeOrgChartProps } from "./ITreeOrgChartProps";
import { ITreeOrgChartState } from "./ITreeOrgChartState";
import "react-sortable-tree/style.css";
import {IPersonaSharedProps, Persona,PersonaSize,} from "office-ui-fabric-react/lib/Persona";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import SPService from "../../../services/SPServices";
import { ITreeChildren } from "./ITreeChildren";
import { ITreeData } from "./ITreeData";
import {Spinner,SpinnerSize} from "office-ui-fabric-react/lib/components/Spinner";
import { Web } from "sp-pnp-js";
import {PeoplePicker,PrincipalType} from"@pnp/spfx-controls-react/lib/PeoplePicker";
import { IconButton } from "office-ui-fabric-react/lib/Button";
import SortableTree from "react-sortable-tree";
//import OrgChart from"@dabeng/react-orgchart"

import Card from '@material-ui/core/Card';
import CardContent from '@material-ui/core/CardContent';
import CardMedia from '@material-ui/core/CardMedia';
import Typography from '@material-ui/core/Typography';
import CardActionArea from '@material-ui/core/CardActionArea';
import Button from '@material-ui/core/Button';
import CardActions from '@material-ui/core/CardActions';

export default class TreeOrgChart extends React.Component<
  ITreeOrgChartProps,
  ITreeOrgChartState
> {
  private treeData: ITreeData[];
  private SPService: SPService;

  constructor(props) {
    super(props);

    this.SPService = new SPService(this.props.context);
    this.state = {
      treeData: [],
      isLoading: true,
      userEmail:'',
      userId:0,
      
    };
  }
  private handleTreeOnChange(treeData) {
    this.setState({ treeData });
  }


  public getUserId(email: string): Promise<number> {
    const web: Web = new Web(this.props.customUrl);
    return web.ensureUser(email).then(result => {
    return result.data.Id;
    });}

  public _getPeoplePickerUserItems = (items: any[]) => {
    if (items.length > 0) {
      var userEmail = items[0].secondaryText;
      this.getUserId(userEmail).then(userId => {
      this.setState({
      userEmail: userEmail,
      userId: userId
    },()=>{this.loadOrgchart(userEmail)});
    });
    }
    else{
      this.setState({
      userEmail:"",
      userId:0
    });
    }
    }

  public async componentDidUpdate(
    prevProps: ITreeOrgChartProps,
    prevState: ITreeOrgChartState
  ) {
    if (
      this.props.currentUserTeam !== prevProps.currentUserTeam ||
      this.props.maxLevels !== prevProps.maxLevels
    ) {
      await this.loadOrgchart(this.props.context.pageContext.user.loginName);
    }
  }


  public async componentDidMount() {
    await this.loadOrgchart(this.props.context.pageContext.user.loginName);
  }

  
  public async loadOrgchart(newValue) {
    this.setState({ treeData: [], isLoading: true });
    const currentUser = `i:0#.f|membership|${newValue}`;
    const currentUserProperties = await this.SPService.getUserProperties(
      currentUser
    );
    
    this.treeData = [];
    if (!this.props.currentUserTeam) {
      const treeManagers = await this.buildOrganizationChart(
        currentUserProperties
      );
      if (treeManagers) this.treeData.push(treeManagers);
    } else {
      const treeManagers = await this.buildMyTeamOrganizationChart(
        currentUserProperties
      );
      if (treeManagers)
        this.treeData.push({
          
          title: treeManagers.person,
          expanded: true,
          children: treeManagers.treeChildren
        });
    }
    this.setState({ treeData: this.treeData, isLoading: false });
  }

 
  public async buildOrganizationChart(currentUserProperties: any) {
    // Get Managers
    let treeManagers: ITreeData | null = null;
    if (
      currentUserProperties.ExtendedManagers &&
      currentUserProperties.ExtendedManagers.length > 0
    ) {
      treeManagers = await this.getUsers(
        currentUserProperties.ExtendedManagers[0]
      );
    }
    return treeManagers;
  }
  
  private async getUsers(manager: string) {
    let person: any;
    let spUser: IPersonaSharedProps = {};
    const managerProperties = await this.SPService.getUserProperties(manager);
    const imageInitials: string[] = managerProperties.DisplayName.split(" ");

    spUser.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${managerProperties.Email}`;
    spUser.imageInitials = `${imageInitials[0]
      .substring(0, 1)
      .toUpperCase()}${imageInitials[1].substring(0, 1).toUpperCase()}`;
    spUser.text = managerProperties.DisplayName;
    spUser.tertiaryText = managerProperties.Email;
    spUser.secondaryText = managerProperties.Title;
      
    const classNames4 = mergeStyleSets({
      compactCard: {
        color:"black",
        fontWeight:"bold",
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
      },
     
      item: {
        selectors: {
          '&:hover': {
            textDecoration: 'underline',
            cursor: 'pointer',
          },
        },
      },
    });
    
    const onRenderCompactCard4= (): JSX.Element => {
      return (
        <div className={classNames4.compactCard}>
          {managerProperties.DisplayName}<br/>
          {managerProperties.Title}
        </div>
      );
    };
    
    const expandingCardProps4: IExpandingCardProps = {
      onRenderCompactCard: onRenderCompactCard4,
      
    };

   
     person = (
      <HoverCard  expandingCardProps={expandingCardProps4} instantOpenOnClick={true} >
      <Persona
        {...spUser}
        hidePersonaDetails={false}
        size={PersonaSize.size40}
      />   
      </HoverCard>
    );  
    
    if (
      managerProperties.DirectReports &&
      managerProperties.DirectReports.length > 0
    ) {
      const usersDirectReports: any[] = await this.getChildren(
        managerProperties.DirectReports
      );
      
      return { id:1,title: person, expanded: true, children: usersDirectReports };
     
    } else {
      
      return { id:0,title: person };
    }
  }
  
  private async getChildren(userDirectReports: any[]) {
    let treeChildren: ITreeChildren[] = [];
    let spUser: IPersonaSharedProps = {};

    for (const user of userDirectReports) {
      const managerProperties = await this.SPService.getUserProperties(user);
      const imageInitials: string[] = managerProperties.DisplayName.split(" ");
      spUser.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${managerProperties.Email}`;
      spUser.imageInitials = `${imageInitials[0]
        .substring(0, 1)
        .toUpperCase()}${imageInitials[1].substring(0, 1).toUpperCase()}`;
      spUser.text = managerProperties.DisplayName;
      spUser.tertiaryText = managerProperties.Email;
      spUser.secondaryText = managerProperties.Title;
      let send_email_report=`mailto:${managerProperties.Email}`;


      const classNames0 = mergeStyleSets({
        compactCard: {
          color:"black",
          fontWeight:"bold",
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          height: '100%',
        },
        expandedCard: {
          font:"Times New Roman",
          color:"black",
          fontWeight:"lighter",
          marginTop:"20px", 
          alignItems: 'center',
          justifyContent: 'center',
          marginLeft:"100px",   
        },
        btn:{
          marginLeft:"45px",
          width:"130px"   
        },
        lnk:{
          marginLeft:"10px",
          fontFamily:"bold"
        },
        root: {
          height:395,
          maxWidth: 345,
         
        },
        media: {
          height: 200,
          
        },
        item: {
          selectors: {
            '&:hover': {
              textDecoration: 'underline',
              cursor: 'pointer',
            },
          },
        },
      });
      
      const onRenderCompactCard0 = (): JSX.Element => {
        return (
          <div className={classNames0.compactCard}>
            {managerProperties.DisplayName}<br/>
            {managerProperties.Title}
          </div>
        );
      };

      const onRenderExpandedCard0 = (): JSX.Element => {
        return (
   <Card  className={classNames0.root}>
      <CardActionArea >
        <CardMedia
          className={classNames0.media}
          image={`/_layouts/15/userphoto.aspx?size=L&username=${managerProperties.Email}`}
          title={managerProperties.DisplayName}
        />
        <CardContent>
          <Typography gutterBottom variant="h5" component="h2">
          {managerProperties.DisplayName}
          </Typography>
          <Typography variant="body2" color="textSecondary" component="p">
          {managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length>0 && managerProperties.UserProfileProperties.find(x=>x.Key=='UserName')?managerProperties.UserProfileProperties.find(x=>x.Key=='UserName').Value?(<span>{managerProperties.UserProfileProperties.find(x=>x.Key=='UserName').Value}<br/></span>):null:null}
          {managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length>0 && managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone')?managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value?(<span>{managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value}<br/></span>):null:null}
          {managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length>0 && managerProperties.UserProfileProperties.find(x=>x.Key=='Office')?managerProperties.UserProfileProperties.find(x=>x.Key=='Office').Value?(<span>{managerProperties.UserProfileProperties.find(x=>x.Key=='Office').Value}<br/></span>):null:null}
          </Typography>
        </CardContent>
      </CardActionArea>
      <CardActions>
        <Button onClick={()=>this.loadOrgchart(managerProperties.Email)}  size="small" color="primary">
          Visit OrgChart
        </Button>
        <Button href={send_email_report} size="small" color="primary">
          Send Email
        </Button>
        <Button href={managerProperties.UserUrl} size="small" color="primary">
          Sharepoint 
        </Button>
      </CardActions>
    </Card>
        );
      };
      const expandingCardProps0: IExpandingCardProps = {
        onRenderCompactCard: onRenderCompactCard0,
        onRenderExpandedCard:onRenderExpandedCard0,
        expandedCardHeight:395
        
      };  
      const person = (
        <HoverCard  expandingCardProps={expandingCardProps0} instantOpenOnClick={true}>
        <Persona
          {...spUser}
          hidePersonaDetails={false}
          size={PersonaSize.size40}
        />
        </HoverCard>
      );
      const usersDirectReports = await this.getChildren(
        managerProperties.DirectReports
      );

      usersDirectReports
        ? treeChildren.push({ title: person, children: usersDirectReports })
        : treeChildren.push({ title: person });
    }
    return treeChildren;
  }

  
  private async buildMyTeamOrganizationChart(currentUserProperties: any) {
    let manager: IPersonaSharedProps = {};
    let me: IPersonaSharedProps = {};
    let treeChildren: ITreeChildren[] = [];
    let imageInitials: string[];
    let hasManager: boolean = false;
    let managerCard: any;

    
    const myManager = await this.SPService.getUserProfileProperty(
      currentUserProperties.AccountName,
      "Manager"
    );


    
    if (myManager) {
      const managerProperties = await this.SPService.getUserProperties(
        myManager
      );
      
      imageInitials = managerProperties.DisplayName?.split(" ").map(name => name[0]);
      manager.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${managerProperties.Email}`;
      
      if (imageInitials)
      manager.imageInitials = `${imageInitials[0]}${imageInitials[1]}`.toUpperCase();
      manager.text = managerProperties.DisplayName;
      manager.tertiaryText = managerProperties.Email;
      manager.secondaryText = managerProperties.Title;
      let send_email_manager=`mailto:${managerProperties.Email}`;

      
      
      const classNames = mergeStyleSets({
        compactCard: {
          color:"black",
          fontWeight:"bold",
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          height: '100%'
        },
        expandedCard: {
          font:"Times New Roman",
          color:"black",
          fontWeight:"lighter",
          marginTop:"20px", 
          alignItems: 'center',
          justifyContent: 'center',
          marginLeft:"100px",
              
        },
        root: {
          height:395,
          maxWidth: 345,
          
        },
        media: {
          height: 200,
          
        },
        
        item: {
          selectors: {
            '&:hover': {
              textDecoration: 'underline',
              cursor: 'pointer',
            },
          },
        },
      });
      
      const onRenderCompactCard = (): JSX.Element => {
        return (
          <div className={classNames.compactCard}>
            {managerProperties.DisplayName}<br/>
            {managerProperties.Title}
          </div>
        );
      };
      const onRenderExpandedCard = (): JSX.Element => {
      
        return (
          <Card className={classNames.root}>
      <CardActionArea>
        <CardMedia
          className={classNames.media}
          image={manager.imageUrl}
          title={managerProperties.DisplayName}
        />
        <CardContent>
          <Typography gutterBottom variant="h5" component="h2">
          {managerProperties.DisplayName}
          </Typography>
          <Typography variant="body2" color="textSecondary" component="p">
          {managerProperties.Email}<br/>
          {managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length>0 && managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone')?managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value?(<span>{managerProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value}<br/></span>):null:null}
            {managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length>0 && managerProperties.UserProfileProperties.find(x=>x.Key=='Office')?managerProperties.UserProfileProperties.find(x=>x.Key=='Office').Value?(<span>{managerProperties.UserProfileProperties.find(x=>x.Key=='Office').Value}<br/></span>):null:null}
          </Typography>
        </CardContent>
      </CardActionArea>
      <CardActions>
        <Button onClick={()=>this.loadOrgchart(managerProperties.Email)}  size="small" color="primary">
          Visit OrgChart
        </Button>
        <Button href={send_email_manager} size="small" color="primary">
          Send Email
        </Button>
        <Button href={managerProperties.UserUrl} size="small" color="primary">
          Visit Sharepoint 
        </Button>
      </CardActions>
    </Card>
          
        );
      };
      const expandingCardProps: IExpandingCardProps = {
        onRenderCompactCard: onRenderCompactCard,
        onRenderExpandedCard:onRenderExpandedCard,
        expandedCardHeight:395
      };
      managerCard = (
        <HoverCard  expandingCardProps={expandingCardProps} instantOpenOnClick={true} >
        <Persona
          {...manager}
          size={PersonaSize.size48}
           coinSize={60} 
          hidePersonaDetails={false}
          
        />
        </HoverCard>
      );
      hasManager = true;
    }

  

    const meImageInitials: string[] = currentUserProperties.DisplayName.split(" ");
    me.imageUrl = `/_layouts/15/userphoto.aspx?size=L&username=${currentUserProperties.Email}`;
    me.imageInitials = `${meImageInitials[0]
      .substring(0, 1)
      .toUpperCase()}${meImageInitials[1].substring(0, 1).toUpperCase()}`;
    me.text = currentUserProperties.DisplayName;
    me.tertiaryText = currentUserProperties.Email;
    me.secondaryText = currentUserProperties.Title;
    let send_email_report=`mailto:${currentUserProperties.Email}`;
    
  
    const classNames2 = mergeStyleSets({
      compactCard: {
        textShadow:"100",
        color:"black",
        fontWeight:"bold",
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
      },

      person:{
      },
      
      expandedCard: {
        marginTop:20,
        font:"Times New Roman",
        color:"black",
        fontWeight:"lighter", 
        alignItems: 'center',
        justifyContent: 'center',
        marginLeft:"50px", 
      },
      btn:{
        padding:"20px",
        marginLeft:"45px",
        
        
      },
      lnk:{
        padding:"20px",
        marginLeft:"10px",
        fontFamily:"bold"
      },
      root: {
        height:395,
        maxWidth: 345,
        
      },
      media: {
        height: 200,
        
      },
      
      item: {
        selectors: {
           '&:hover': {
            
            textDecoration: 'underline',
            cursor: 'pointer',
            alignContent:"center"
          },
        },
      },
    });
  
    const onRenderExpandedCard2 = (): JSX.Element => {
      return (
        <Card className={classNames2.root}>
      <CardActionArea>
        <CardMedia
          className={classNames2.media}
          image={`/_layouts/15/userphoto.aspx?size=L&username=${currentUserProperties.Email}`}
          title={currentUserProperties.DisplayName}
        />
        <CardContent>
          <Typography gutterBottom variant="h5" component="h2">
          {currentUserProperties.DisplayName}
          </Typography>
          <Typography variant="body2" color="textSecondary" component="p">
          {currentUserProperties.Email}<br/>
          {currentUserProperties && currentUserProperties.UserProfileProperties && currentUserProperties.UserProfileProperties.length>0 && currentUserProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone')?currentUserProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value?(<span>{currentUserProperties.UserProfileProperties.find(x=>x.Key=='WorkPhone').Value}<br/></span>):null:null}
          {currentUserProperties && currentUserProperties.UserProfileProperties &&currentUserProperties.UserProfileProperties.length>0 && currentUserProperties.UserProfileProperties.find(x=>x.Key=='Office')?currentUserProperties.UserProfileProperties.find(x=>x.Key=='Office').Value?(<span>{currentUserProperties.UserProfileProperties.find(x=>x.Key=='Office').Value}<br/></span>):null:null}
          </Typography>
        </CardContent>
      </CardActionArea>
      <CardActions>
        <Button onClick={()=>this.loadOrgchart(currentUserProperties.Email)}  size="small" color="primary">
          Visit OrgChart
        </Button>
        <Button href={send_email_report}  size="small" color="primary">
          Send Email
        </Button>
        <Button href={currentUserProperties.UserUrl}  size="small" color="primary">
          Visit Sharepoint 
        </Button>
      </CardActions>
    </Card>
        
 

      );
    };

    const onRenderCompactCard2 = (): JSX.Element => {
      return (
        <div className={classNames2.compactCard}>
          {me.text}<br/>
          {me.secondaryText}<br/>
        </div>
      );
    };

    const expandingCardProps2: IExpandingCardProps = {
      onRenderCompactCard: onRenderCompactCard2,
      onRenderExpandedCard:onRenderExpandedCard2,
      expandedCardHeight:395
    };
  
  
    
    const meCard = (
     <div> 
        <HoverCard  expandingCardProps={expandingCardProps2} instantOpenOnClick={true} >
      <Persona {...me} initialsColor="blue"   className={classNames2.person}  hidePersonaDetails={false} size={PersonaSize.size48} coinSize={60}  />
     </HoverCard>

    </div>
    );
    const usersDirectReports: any[] = await this.getChildren(
      currentUserProperties.DirectReports
    );
   
    if (hasManager) {
      treeChildren.push({
        
        title: meCard,
        expanded: true,
        children: usersDirectReports
      });
    } else {
      treeChildren = usersDirectReports;
      managerCard = meCard;
    }
    return { person: managerCard, treeChildren: treeChildren };
    } 
    
  public render(): React.ReactElement<ITreeOrgChartProps> {

  
    return(
        <div className={styles.treeOrgChart}> 
        
          <PeoplePicker
            context={this.props.context}
            titleText=""
            personSelectionLimit={1}
            showtooltip={true}
            defaultSelectedUsers={
            this.state.userEmail ? [this.state.userEmail] : []}
            selectedItems={this._getPeoplePickerUserItems.bind(this)}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
      
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
        />
        {this.state.isLoading ? (
          <Spinner
            size={SpinnerSize.large}
            label="Loading Organization Chart ..."
          ></Spinner>
        ) : null}

        <div className={styles.treeContainer}>
        <SortableTree
            treeData={this.state.treeData}
            onChange={this.handleTreeOnChange.bind(this)}
            canDrag={false}
            canDrop={false}
            rowHeight={120}
            scaffoldBlockPxWidth={100}
            rowDirection="ltr"
            orientation="horizontal"
            maxDepth={this.props.maxLevels}
            generateNodeProps={rowInfo => ({
              buttons: [ 
                
                <IconButton
                  disabled={false}
                  checked={false}
                  size={60}
                  iconProps={{ iconName: "ContactInfo" }}
                  title="Contact Info"
                  ariaLabel="Contact"
                  onClick={() => {
                    window.open(
                      `https://nam.delve.office.com/?p=${rowInfo.node.title.props.children.props.tertiaryText}&v=work`
                    );
                      
                    
                    
                  }}
                />
                
              ]
            })}
          />
       
       
       
        </div>
      </div>
    );
  }
}

