import { sp } from "@pnp/sp/presets/all";

export class SPService {
  private listItem: any;
  private listTitle: string;

  constructor(listTitle: string, item: any) {
    this.listItem = item;
    this.listTitle = listTitle;
  }

  /**
   * Returns the item by its Id from the specified list.
   */
  public getItem<T>(): Promise<T> {
    return new Promise((resolve, reject) => {
      sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(parseInt(this.listItem.id.toString()))
        .get<T>()
        .then((item: T) => {
          resolve(item);
        })
        .catch((error) => {
          reject(error);
        });
    });
  }
}
