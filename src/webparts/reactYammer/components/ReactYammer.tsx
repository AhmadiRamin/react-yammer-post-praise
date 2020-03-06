import * as React from 'react';
import styles from './ReactYammer.module.scss';
import { IReactYammerProps } from './IReactYammerProps';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import LoadingImage from './loading';
import IPraise from '../interface/IPraise';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import badges from './badges';

const ReactYammer: React.SFC<IReactYammerProps> = (props) => {

  const [loading, setLoading] = React.useState(false);
  const [formVisible, setFormVisible] = React.useState(true);
  const [nominee, setNominee] = React.useState("");
  const [icon, setIcon] = React.useState("FavoriteStar");
  const [groups,setGroups] = React.useState<IDropdownOption[]>([]);
  const [headline, setHeadline] = React.useState("");
  const [praise, setPraise] = React.useState("");
  const [groupId, setGroupId] = React.useState("");
  const [messageBarStatus, setMessageBarStatus] = React.useState({
    type: MessageBarType.info,
    message: '',
    show: false
  });


  React.useEffect(()=>{
    props.yammerProvider.getGroups().then(grps=>{
     
      const options: IDropdownOption[] = grps.data.map(g=>  ({key:g.id,text:g.full_name}));
      console.log(options);
      setGroups(options);
    }).catch(err=>{
      console.log(err);
    });
  },[]);

  const _getPeoplePickerItems = (items: any[]) => {
    setNominee(items[0].secondaryText);
  };

  const _onIconChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setIcon(item.text);
  };

  const _onGroupChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setGroupId(item.key.toString());
  };

  const _postPraise = () => {
    const objPraise: IPraise = {
      icon,
      headline,
      nominee,
      praise,
      groupId
    };
    setLoading(true);
    props.yammerProvider.postPraise(objPraise).then(response => {
      setMessageBarStatus({
        type: MessageBarType.success,
        message: "Your priase now been successfully added.",
        show: true
      });
      setLoading(false);

    }).catch(error => {
      setMessageBarStatus({
        type: MessageBarType.error,
        message: "Unfortunately we could not post your praise. Please try again later.",
        show: true
      });
      setLoading(false);
    });
  };

  return (
    <div className={styles.reactYammer}>
        <div>
          {
            loading && <LoadingImage />
          }
        </div>
        <div>
          {
            messageBarStatus.show &&
            <MessageBar messageBarType={messageBarStatus.type}>{
              messageBarStatus.message
            }</MessageBar>
          }
        </div>
        {
          (!loading && formVisible) &&
          <div>
            <div>
              <PeoplePicker isRequired
                context={props.context}
                titleText="Nominee"
                ensureUser={true}
                personSelectionLimit={1}
                groupName=""
                selectedItems={_getPeoplePickerItems}
                defaultSelectedUsers={[nominee]}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
              
              <TextField required placeholder="Please enter the Headline" label="Headline" value={headline} onChanged={(value) => setHeadline(value)} />
              
              <TextField required maxLength={250} placeholder="Describe what they've done." label="Reason" value={praise} multiline={true} rows={6} onChanged={(value) => setPraise(value)} />
              
              <Dropdown required label="Group"
                options={groups} onChange={_onGroupChange}/>

              <Dropdown required label="Icon"
                options={badges} onChange={_onIconChange} />
            </div>
            <br />
            <div title="Please fill in all required fields.">
              <DefaultButton text="Post" title="Please fill in all required fields" onClick={_postPraise} disabled={headline === "" || praise === "" || icon === "" || nominee === "" || groupId ===""} />
            </div>
          </div>
        }
    </div>
  );
};

export default ReactYammer;