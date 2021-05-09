import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import styles from './SpFxDemo.module.scss';
import { ISpFxDemoProps } from './ISpFxDemoProps';
import SPService from '../services/SPService';
import { ShoppingList } from './ShoppingList';

export interface ISPFxDemoState {
  items: any[];
}

export default class SpFxDemo extends React.Component<ISpFxDemoProps, ISPFxDemoState> {

  service: SPService;


  constructor(props: ISpFxDemoProps) {
    super(props);
    this.state = {
      items: []
    }
    this.service = new SPService(this.props.context);
  }

  componentDidMount() {
    if (this.service) {
      this.service.getBirthdayItems()
        .then(
          response => {
            console.log('SharePoint Data is ', response.value);
            let spItems = [];

            for (let item of response.value) {
              let bDay = new Date(item.BirthDate);
              let today = new Date();
              if (bDay.getMonth() == today.getMonth())
                spItems.push(item);
            }
            console.log('SP Items ', spItems);
            this.setState({ items: spItems });
          },
          error => console.error('Oops error', error)
        )
    }
  }



  public render(): React.ReactElement<ISpFxDemoProps> {
    let persons: IPersonaSharedProps[] = this.state.items.map(p => {

      let photo: string = `/_layouts/15/userphoto.aspx?size=L&username=${p.Employee.EMail}`;
      let f = Intl.DateTimeFormat('en-us', { month: 'long', day: 'numeric' });

      return {
        imageUrl: photo,
        imageAlt: p.Employee.Title + ' Photo',
        text: p.Employee.Title,
        secondaryText: p.Employee.JobTitle,
        tertiaryText: f.format(new Date(p.BirthDate)),
        optionalText: 'Department: ' + p.Employee.Department
      }


    });
    return (
      <div className={styles.spFxDemo}>
        <h1>{this.props.description}</h1>
        <ul>
          {persons.map(item => { 
            return <>
              <Persona
                  {...item}
                  size={PersonaSize.size100}
                  presence={PersonaPresence.online}
                  hidePersonaDetails={false}
                />
                <hr/>
            </> 
            })}
        </ul>
        <ShoppingList name="Tesfaye List"></ShoppingList>
      </div>
    );
  }
}
