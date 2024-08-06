import { IPokemonItem } from "./IPokemonItem";
import { IPokemonType } from "./IPokemonType";

export interface IGraphSearchState {
  pokemons: Array<IPokemonItem>;
  types: Array<IPokemonType>;
  typeValue: string;
  searchFor: {
    search:string,
    type:string
  };
  page: number;
  totalPage: number;
  isLoading: boolean;
  togglePanel: boolean;
}