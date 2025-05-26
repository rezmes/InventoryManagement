// src\webparts\inventory\utils\findOption.ts
   import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown'

   export function findOption(options: IDropdownOption[], key: string | number | undefined): IDropdownOption | undefined {
     if (key === undefined) return undefined
     for (var i = 0; i < options.length; i++) {
       if (options[i].key === key) {
         return options[i]
       }
     }
     return undefined
   }
