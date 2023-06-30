import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import Category from "./Category";
import MailBodyUpdator from "./MailBodyUpdator";
import SupportedVersion from "./SupportedVersion";
import MultipleSelect from "./MultipleSelect";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        // {
        //   icon: "Ribbon",
        //   primaryText: "Achieve more with Office integration",
        // },
        // {
        //   icon: "Unlock",
        //   primaryText: "Unlock features and functionality",
        // },
        // {
        //   icon: "Design",
        //   primaryText: "Create and visualize like a pro",
        // },
      ],
    });
  }

  click = async () => {
    /**
     * Insert your Outlook code here
     */
    console.log("Is 1.3 supported?"+Office.context.requirements.isSetSupported("Mailbox", '1.3'));
    
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" /> */}
        <HeroList message="CS791 outlook add-on demo" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
        </HeroList>
        <Category />
        <MultipleSelect/>
        <MailBodyUpdator MailBody="This is the hardcoded content"/>
        <SupportedVersion/>
      </div>
    );
  }
}
