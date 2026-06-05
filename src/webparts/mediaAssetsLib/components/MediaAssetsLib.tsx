import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import type { IMediaAssetsLibProps } from './IMediaAssetsLibProps';

interface IMediaItem {
  id: number;
  title: string;
  fileRef: string;
}

interface IMediaAssetsLibState {
  mediaItems: IMediaItem[];
  searchText: string;
}

export default class MediaAssetsLib extends React.Component<
  IMediaAssetsLibProps,
  IMediaAssetsLibState
> {

  constructor(props: IMediaAssetsLibProps) {
    super(props);

    this.state = {
      mediaItems: [],
      searchText: ''
    };
  }

  
  public async componentDidMount(): Promise<void> {
    await this.loadMedia();
  }
  

  /*

  public async componentDidMount(): Promise<void> {

  this.setState({
    mediaItems: [
      {
        id: 1,
        title: 'Testbild',
        fileRef: '/sites/Medienbibliothek/Medien/test.png'
      }
    ]
  });

}

*/

  
  private async loadMedia(): Promise<void> {

    const url =
      `${this.props.siteUrl}/_api/web/lists/getbytitle('Videos')/items` +
      `?$select=Id,Title,FileRef,Eigene_x0020_Tags`;

    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    const data = await response.json();

    const mediaItems = data.value.map((item: any) => ({
      id: item.Id,
      title: item.Title,
      fileRef: item.FileRef,
      tags: item.Eigene_x0020_Tags
    }));

    this.setState({
      mediaItems
    });
  }
  

  private onSearchChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ): void => {

    this.setState({
      searchText: event.target.value
    });
  };

  public render(): React.ReactElement<IMediaAssetsLibProps> {

    const search = this.state.searchText.toLowerCase();

    const filtered = this.state.mediaItems.filter(item =>
      item.title.toLowerCase().indexOf(search) > -1
    );

    return (
      <div style={{ padding: '20px' }}>

        <h2>Media Assets Library</h2>

        <input
          type="text"
          placeholder="Suche..."
          value={this.state.searchText}
          onChange={this.onSearchChange}
        />

        <div
          style={{
            display: 'flex',
            flexWrap: 'wrap',
            gap: '20px',
            marginTop: '20px'
          }}
        >

          {filtered.map(item => (

            <div
              key={item.id}
              style={{
                width: '300px',
                border: '1px solid #ddd',
                padding: '10px'
              }}
            >

              <img
                src={`${window.location.origin}${item.fileRef}`}
                alt={item.title}
                width="100%"
              />

              <h3>{item.title}</h3>

            </div>

          ))}

        </div>

      </div>
    );
  }
}