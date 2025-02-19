// import * as React from 'react';
// import { useEffect } from 'react';
// import { Modal, IconButton, IIconProps, getTheme, mergeStyleSets, FontWeights } from '@fluentui/react';
// import { IDataModalProps } from '../interfaces';

// const cancelIcon: IIconProps = { iconName: 'Cancel' };

// const theme = getTheme();
// const contentStyles = mergeStyleSets({
//   container: {
//     display: 'flex',
//     flexDirection: 'column',
//     alignItems: 'center',
//     padding: '20px',
//   },
//   header: [
//     theme.fonts.xLarge,
//     {
//       fontWeight: FontWeights.semibold,
//       marginBottom: '20px',
//     },
//   ],
//   body: {
//     marginBottom: '20px',
//   },
// });

// const DataModal: React.FC<IDataModalProps> = ({ isOpen, modalData, onDismiss }) => {
//   const [groupData, setGroupData] = React.useState(modalData);

//   useEffect(() => {
//     setGroupData(modalData);
//   }, [modalData]);

//   return (
//     <Modal
//       isOpen={isOpen}
//       onDismiss={onDismiss}
//       isBlocking={false}
//       containerClassName={contentStyles.container}
//     >
//       <div className={contentStyles.header}>
//         <span>{"My Groups"}</span>
//         <IconButton
//           styles={{ root: { marginLeft: 'auto' } }}
//           iconProps={cancelIcon}
//           ariaLabel="Close popup modal"
//           onClick={onDismiss}
//         />
//       </div>
//       <div className={contentStyles.body}>
//         <ul>
//           {groupData.map((item, index) => (
//             <li key={index}>{item.Title}</li>
//           ))}
//         </ul>
//       </div>
//     </Modal>
//   );
// };

// export default DataModal;