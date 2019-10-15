//Tiles.tsx

import * as React from 'react';
import styles from './Tiles.module.scss';
import { ITilesProps } from './ITilesProps';
import * as pnp from "sp-pnp-js/lib/pnp";
import { Spinner, SpinnerType } from "office-ui-fabric-react/lib/Spinner";

export interface ITilesState {
  items?: Array<any>;
  isLoading?: boolean;
}

export default class Tiles extends React.Component<ITilesProps, ITilesState> {
  constructor(props) {
    super(props);
    this.state = {
      items: [],
      isLoading: true,
    };
  }
  public componentDidMount(): void {
    this.fetchData();
  }
  public componentDidUpdate(prevprops): void {
    if (this.props !== prevprops) {
      this.fetchData();
    }
  }
  public render(): React.ReactElement<ITilesProps> {
    let { isLoading, items }: ITilesState = this.state;
    let elements = items.map((item: any, index: number) => {

      let thisTop = `${this.props.imageHeight / 3 * 2}px`;
      let thisHeight = `${this.props.imageHeight}px`;
      let thisWidth = `${this.props.imageWidth}px`;
      let imgURL = (item[this.props.backgroundImageField]) ? item[this.props.backgroundImageField].Url : this.props.fallbackImageUrl;
      let thisTarget = (item[this.props.newTabField]) ? "_blank" : "";
      let thisHref = (item[this.props.linkField]) ? item[this.props.linkField].Url : "#";
      let thisPadding = `${this.props.textPadding}px`;

      return <a className={styles.promotedLink} style={{ width: thisWidth, height: thisHeight }} key={index} target={ thisTarget } href={ thisHref }>
        <img className={styles.image} src={ imgURL  } />
        <div className={styles.textArea} style={{ height: thisHeight, top: thisTop }}>
          <div className={styles.container} style={{ padding: thisPadding }}>
            <div className={styles.title}>{item.Title}</div>
            <div className={styles.description}>{item[this.props.descriptionField]}</div>
          </div>
        </div>
      </a>;
    });
    if (isLoading) {
      return <Spinner type={SpinnerType.large} />;
    } else {
      return <div>{
        elements.length > 0 && <div className={styles.promotedLinks}>
          {elements}
        </div>}
      </div>;

    }
  }
  private async fetchData(): Promise<void> {
    try {
      let filter = (this.props.tileTypeField && this.props.tileType) ? `${this.props.tileTypeField} eq '${this.props.tileType}'` : '';
      let response = await pnp.sp.web.lists.getByTitle(this.props.list).items.filter(filter).orderBy((this.props.orderByField) ? this.props.orderByField : "ID").top(this.props.count).get();
      console.log('fetchData():');
      console.log(response);      
      this.setState({
        items: response,
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        isLoading: false,
      });
      throw error;
    }
  }
}
