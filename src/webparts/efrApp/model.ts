import pnp,
{
    SharePointQueryable,
    Item,
   
} from "sp-pnp-js";
export class PBCTask{
    public Id: number;
    public Library:string;
    public kleigfa:string;
    public Title: string;
    public ProductNumber: string;
    public OrderDate: Date;
    public OrderAmount: number;
}