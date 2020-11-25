import { MessageBar,MessageBarType } from 'office-ui-fabric-react';
import * as React from 'react'
import styles from './msg.module.scss';
interface MsgProps{
  show: boolean
}

export class Message extends React.Component<MsgProps,{}>{
  state= {
    show: this.props.show
  }

  public render(): React.ReactElement{
    return(
      <>
        {
          this.state.show ?
            <MessageBar
            className={styles.msg}
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => this.setState({show: false})}
            dismissButtonAriaLabel="Close"
            >
            Selezionare almeno una riga per eseguire l'export.
            </MessageBar>
          : null
        }
      </>
    )
  }
}
