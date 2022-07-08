/* eslint-disable react/jsx-no-bind */
import * as React from "react";
import styles from "./SampleDemo.module.scss";
import { ISampleDemoProps } from "./ISampleDemoProps";

//import library
import { PrimaryButton, Stack, MessageBar } from "office-ui-fabric-react";
import { sp, IItemAddResult } from "@pnp/sp/presets/all";

//create state
export interface ISampleDemoState {
  showmessageBar: boolean;
  message: string;
  itemID: number;
}

export default class SampleDemo extends React.Component<
  ISampleDemoProps,
  ISampleDemoState
> {
  public constructor(props: ISampleDemoProps, state: ISampleDemoState) {
    super(props);
    this.state = { itemID: 0, showmessageBar: false, message: "" };
  }

  public render(): React.ReactElement<ISampleDemoProps> {
    return (
      <div className={styles.sampleDemo}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>
                Welcome to PnP JS List Items Operations Demo!
              </span>
            </div>
          </div>
        </div>
        <br />
        <br />
        <Stack horizontal tokens={{ childrenGap: 40 }}>
          <PrimaryButton
            text="Create New Item"
            onClick={() => this._createNewItem()}
          />
          <PrimaryButton text="Get Item" onClick={() => this._getItem()} />
          <PrimaryButton
            text="Update Item"
            onClick={() => this._updateItem()}
          />
          <PrimaryButton text="Delete Item" onClick={() => this._delteItem()} />
        </Stack>
        <br />
        <br />
        {this.state.showmessageBar && (
          <MessageBar
            onDismiss={() => this.setState({ showmessageBar: false })}
            dismissButtonAriaLabel="Close"
          >
            {this.state.message}
          </MessageBar>
        )}
      </div>
    );
  }

  private _createNewItem(): void {
    sp.web.lists
      .getByTitle("DemoList")
      .items.add({
        Title: "Title " + new Date(),
        Description: "This is item created using PnP JS",
      })
      .then((item: IItemAddResult) => {
        console.info(item);
        this.setState({
          ...this.state,
          itemID: item.data.Id,
          showmessageBar: true,
          message: "Item created successfully",
        });
      })
      .catch(console.error);
  }

  private _getItem(): void {
    // get a specific item by id
    sp.web.lists
      .getByTitle("DemoList")
      .items.getById(this.state.itemID)
      .get()
      .then((item) => {
        console.info(item);
        this.setState({
          ...this.state,
          showmessageBar: true,
          message: "Item retrieved successfully",
        });
      })
      .catch(console.error);
  }

  private _updateItem(): void {
    sp.web.lists
      .getByTitle("DemoList")
      .items.getById(this.state.itemID)
      .update({
        Title: "Title " + new Date(),
        Description: "This is item updated using PnP JS",
      })
      .then((list) => {
        console.log(list);

        this.setState({
          ...this.state,
          showmessageBar: true,
          message: "Item updated successfully",
        });
      })
      .catch(console.error);
  }

  private _delteItem(): void {
    sp.web.lists
      .getByTitle("DemoList")
      .items.getById(this.state.itemID)
      .delete()
      .then((list) => {
        console.log(list);
        this.setState({
          ...this.state,
          showmessageBar: true,
          message: "Item deleted successfully",
        });
      })
      .catch(console.error);
  }
}
