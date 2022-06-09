import * as React from 'react';
import { Label } from '@fluentui/react/lib/Label';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { IBasePickerSuggestionsProps , NormalPeoplePicker, ValidationState } from 'office-ui-fabric-react/lib/Pickers';
function SharedWithPeople(props)
{

    return(<div>
       
        <NormalPeoplePicker onResolveSuggestions={props.GetUserDetails}
        /*createGenericItem={()=>{
            return(<div>dshsdsdhsdhshsdhsdh</div>)
        }}
        onValidateInput={()=>{
            return ValidationState.valid;
        }}*/
        //itemLimit={1}
        onChange={(items)=>{
            console.log(items);
            props.update(items);
        }}
        /*onRemoveSuggestion={(items)=>{
            console.log(items);
        }}
       onItemSelected={(items)=>{//which is used to get the selected item
            console.log(items);
            return items;
       }}*/
        defaultSelectedItems={props.peoples}//which is used for selected items
        inputProps={{
            onBlur: (ev: React.FocusEvent<HTMLInputElement>) =>{ },
            onFocus: (ev: React.FocusEvent<HTMLInputElement>) =>{},
            'aria-label': 'People Picker',
          }}
        />
    </div>)
}

export {SharedWithPeople}