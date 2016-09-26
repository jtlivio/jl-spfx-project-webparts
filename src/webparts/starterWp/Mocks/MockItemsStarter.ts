import { ISPListItem } from '../StarterWpWebPart';
export default class MockItemsStarter {

     private static _items: ISPListItem[] = [
       {
         Title: 'Mock Item',
         Id: '1',
         MyMultiText: '<strong>Lorem Ipsum is simply</strong> dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s'
        }
      ];

    public static get(restUrl: string, options?: any): Promise<ISPListItem[]> {
    return new Promise<ISPListItem[]>((resolve) => {
            resolve(MockItemsStarter._items);
        });
    }
}

