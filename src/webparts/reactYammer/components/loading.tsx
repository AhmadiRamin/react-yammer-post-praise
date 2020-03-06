import * as React from 'react';
const loadingImage: any = require("./assets/loading.gif");

const loading = ()=>{
    return(
        <>
            <img src={loadingImage} width={250} />
        </>
    );
};
export default loading;