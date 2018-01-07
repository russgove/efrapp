// export const load = async (o, include?) => {
//     const ctx = o.get_context();
//     if (include) {
//       ctx.load(o, include);
//     } else {
//       ctx.load(o);
//     }
//     await exec(ctx);
//     if (o.getEnumerator) return toArray(o);
//     return o;
//   };
  
//   export const exec = ctx => new Promise(r => 
//     ctx.executeQueryAsync(r, (a, b) => console.log(b.get_message()))
//   );
  
//   export const toArray = spItem => {
//     const res = [];
//     const enumerator = spItem.getEnumerator();
//     while (enumerator.moveNext()) res.push(enumerator.get_current());
//     return res;
//   };