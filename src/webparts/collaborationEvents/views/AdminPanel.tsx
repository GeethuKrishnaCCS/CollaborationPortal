// import * as React from 'react';
// import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
// import { Panel } from '@fluentui/react/lib/Panel';
//import { IAdminPanelProps } from '../interfaces';

//const buttonStyles = { root: { marginRight: 8 } };

// export const AdminPanel: React.FunctionComponent<IAdminPanelProps> = ({ showPanel, onDismiss }) => {
//   // This panel doesn't actually save anything; the buttons are just an example of what
//   // someone might want to render in a panel footer.
//   const onRenderFooterContent = React.useCallback(
//     () => (
//       <div>
//         <PrimaryButton onClick={onDismiss} styles={buttonStyles}>
//           Save
//         </PrimaryButton>
//         <DefaultButton onClick={onDismiss}>Cancel</DefaultButton>
//       </div>
//     ),
//     [onDismiss],
//   );

//   return (
//     <div>
//       <Panel
//         isOpen={showPanel}
//         onDismiss={onDismiss}
//         headerText="Admin Panel"
//         closeButtonAriaLabel="Close"
//         onRenderFooterContent={onRenderFooterContent}
//         isFooterAtBottom={true}
//       >
//         <p>Admin Controls go here.</p>
//       </Panel>
//     </div>
//   );
// };