import { ISPList } from './AccordionWebPart';  

export default class MockHttpClient {  
    private static _items: ISPList[] = [{ Title: 'Heading 1', Content: 'Content 1' },];  
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }  
}  