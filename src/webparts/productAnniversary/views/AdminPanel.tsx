import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
import { useBoolean } from '@fluentui/react-hooks';
import { AdminPanelProps} from '../interfaces';

const buttonStyles = { root: { marginRight: 8 } };

export const AdminPanel: React.FunctionComponent<AdminPanelProps> = ({ showPanel }) => {
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(showPanel);

  // This panel doesn't actually save anything; the buttons are just an example of what
  // someone might want to render in a panel footer.
  const onRenderFooterContent = React.useCallback(
    () => (
      <div>
        <PrimaryButton onClick={dismissPanel} styles={buttonStyles}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={dismissPanel}>Cancel</DefaultButton>
      </div>
    ),
    [dismissPanel],
  );

  React.useEffect(() => {
    if (showPanel) {
      openPanel();
    } else {
      dismissPanel();
    }
  }, [showPanel, openPanel, dismissPanel]);

  return (
    <div>
      <Panel
        isOpen={isOpen}
        onDismiss={dismissPanel}
        headerText="AdminPanel"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={onRenderFooterContent}
        // Stretch panel content to fill the available height so the footer is positioned
        // at the bottom of the page
        isFooterAtBottom={true}
      >
        <p>Admin Controls goes here.</p>
      </Panel>
    </div>
  );
};