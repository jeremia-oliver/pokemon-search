import * as strings from 'GraphSearchWebPartStrings';
import * as React from 'react';
import styles from './GraphSearch.module.scss';
import { escape } from "@microsoft/sp-lodash-subset";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import type { IGraphSearchProps } from './IGraphSearchProps';
import { IGraphSearchState } from "./IGraphSearchState";
import { IPokemonItem } from "./IPokemonItem";
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";
import { BaseButton,Button ,PrimaryButton, DefaultButton, TextField, Dropdown, IDropdownOption, IIconProps } from '@fluentui/react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { IPokemonType } from './IPokemonType';

export default class GraphSearch extends React.Component<IGraphSearchProps,IGraphSearchState> {
  constructor(props: IGraphSearchProps, state: IGraphSearchState) {
    super(props);
    this.state = {pokemons: [], types:[], typeValue:"" ,searchFor: {search:"", type:""}, page: 1, totalPage: 0, isLoading:true, togglePanel: false};
    this._searchWithGraph(1);
  }
  public render(): React.ReactElement<IGraphSearchProps> {  
    const {
      hasTeamsContext,
    } = this.props;
    const closeIcon: IIconProps = { iconName: 'Cancel' };
    return (
      <section className={`${styles.graphSearch} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
            <p className={ styles.form }>
              <TextField
                  label={ strings.SearchFor }
                  required={ true }
                  onChange={ this._onSearchForChanged }
                  onGetErrorMessage={ this._getSearchForErrorMessage }
                  value={ this.state.searchFor.search }
                />
            </p>
            <div className={ styles.filterButtonContainer }>
              <div>
                <PrimaryButton
                    text='Search'
                    title='Search'
                    onClick={ this._search }
                  />
              </div>
              <div>
                <PrimaryButton text="Filter" disabled={this.state.types.length == 0} onClick={this._togglePanel}><Icon iconName="Filter"/></PrimaryButton>
                <Panel
                    isOpen={this.state.togglePanel}
                    onDismiss={this._togglePanel}
                    headerText="Filter Pokemon"
                    closeButtonAriaLabel="Close"
                >
                    <Dropdown
                      placeholder="Select a type"
                      label="Pokemon Type"
                      options={this.state.types}
                      onChange={this._changeType}
                      defaultSelectedKey={this.state.searchFor.type}
                    />
                    <div className={styles.panelFooterContainer}>
                        <PrimaryButton onClick={ this._filterSubmit }>Filter</PrimaryButton>
                        <DefaultButton onClick={ this._togglePanel }>Cancel</DefaultButton>
                    </div>
                </Panel>
              </div>
            </div>
        </div>
        {
        this.state.searchFor.type !== '' ? (
          <div className={styles.TagStylesContainer}>
          <div className={styles.TagStyles}><span>Type: { this.state.searchFor.type }</span> <IconButton iconProps={closeIcon} title="Remove Filter" ariaLabel="Remove Filter" onClick={ this._removeTypeFilter }/></div></div>
        ) : null
        }
        {
        this.state.isLoading ? (
          <div className={styles.loading}>
              <Spinner label="Loading Pokemon..." />
          </div>
        ) : (
          <div className={styles.FlexContainer}>
            {
              this.state.pokemons.length > 0 ?
                this.state.pokemons.map((e) => {
                  return(
                  <div 
                  className={styles.ColContainer}
                  id={e.name}>
                      <div className={styles.ColContainerImage}>
                      <img
                          className={styles.ColImage}
                          src={ e.documentLink }
                          alt={ e.name }
                      />
                      </div>
                      <span className={styles.PokemonId}>#{ e.pokedex }</span>
                      <span>{ e.name }</span>
                  </div>
                  )
                })
              : null
            }
          </div>
        )
        }
        {
        this.state.totalPage > 0 ?
        <Pagination
          currentPage={this.state.page}
          totalPages={this.state.totalPage} 
          onChange={(page) => this._getPage(page)}
          limiter={ 2 } 
        />
        : null
        }
      </section>
    );
  }

  private _togglePanel = ():void => {
    this.setState({
      togglePanel: !this.state.togglePanel
    })
  }
  private _changeType = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({
      typeValue: item.key as string
    })
  }
  private _removeTypeFilter = async() : Promise<void> => {
    await this.setState({
      typeValue: ""
    })
    this._filterSubmit();
  }
  private _filterSubmit = async() : Promise<void> => {
    await this.setState(prevState => ({
      searchFor:{
        ...prevState.searchFor, 
        type: this.state.typeValue
      },
      togglePanel: false,
      isLoading: true
    }));
    this._searchWithGraph(1);
  }
  private _getPage = (thisPage: number): void => {
    this.setState(
      {
        isLoading: true,
      }
    )
    this._searchWithGraph(thisPage);
  }

  private _onSearchForChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    // Update the component state accordingly to the current user's input
    this.setState(prevState => ({
      searchFor:{
        ...prevState.searchFor, 
        search: newValue ? newValue :  ""
      }
    }));
  }
  
  private _getSearchForErrorMessage = (value: string): string => {
    // The search for text cannot contain spaces
    return (value === null || value.length === 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  private _search = (event: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button, MouseEvent>) : void => {
    this.setState(
      {
        isLoading: true,
      }
    )
    this._searchWithGraph(1);
  }
  
  private _searchWithGraph =  async(thisPage: number) : Promise<void> => {
    // Log the current operation
    const filterType = this.state.searchFor.type !== "" ? `ctpokemontype:${this.state.searchFor.type} ` : "";
    const filterSearch = this.state.searchFor.search !== "" ? `${escape(this.state.searchFor.search)}*` : "";
    console.log(this.state.searchFor);
    const Payload = {
      "requests": [
          {
              "entityTypes": [
                  "driveItem"
              ],
              "query": {
                  "queryString": `path:https://ctlab03.sharepoint.com/sites/devjeremia/Pokemon%20Lib ${filterType}${filterSearch}`
              },
              "fields": [
                  "id",
                  "name",
                  "contentclass",
                  "title",
                  "path",
                  "filetype",
                  "listitemid",
                  "spweburl",
                  "spsiteurl",
                  "uniqueid",
                  "DocumentLink",
                  "refinablestring00",
                  "ctpokemontype",
                  "ctpokedexnumber"
              ],
              "from": thisPage > 0 ? (thisPage - 1) * 20 : 0,
              "size": 20,
              "sortProperties": [
                  {
                      "name": "ctpokedexnumber",
                      "isDescending": false
                  }
              ],
              "aggregations": [
                  {
                      "field": "ctpokemontype",
                      "bucketDefinition": {
                          "sortBy": "keyAsString",
                          "isDescending": "false"
                      }
                  }
              ]
          }
      ]
  }
      await this.props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) :void => {
        client
          .api("search/query")
          .post((Payload), (err, res) => {
            if (err) {
              console.error(`Error: ${err}`);
              return;
            }
            const pokemons: Array<IPokemonItem> = new Array<IPokemonItem>();
            const types:Array<IPokemonType> = this.state.types;
            if(res.value[0].hitsContainers[0].hits){
              if(this.state.types.length == 0){
                res.value[0].hitsContainers[0].aggregations.map((el: any) => {
                  if(el.field === 'ctpokemontype'){
                    el.buckets.map((e: any) => {
                      types.push({
                        key: e.key,
                        text: e.key
                      })
                    })
                  }
                })
              }
              res.value[0].hitsContainers[0].hits.map((item: any) => {
                pokemons.push({
                  documentLink: item.resource.listItem.fields.documentLink,
                  pokedex: item.resource.listItem.fields.ctpokedexnumber,
                  name: item.resource.listItem.fields.title,
                  type: item.resource.listItem.fields.ctpokemontype
                })
              });
              this.setState(
                {
                  page: thisPage,
                  pokemons: pokemons,
                  totalPage: Math.ceil(res.value[0].hitsContainers[0].total / 20),
                  isLoading: false,
                  types: types
                }
              );
            }else{
              this._searchWithGraph(thisPage)
            }
          });
      });
  }
}

