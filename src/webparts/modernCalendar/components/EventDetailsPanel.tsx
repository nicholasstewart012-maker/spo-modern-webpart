import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './EventDetailsPanel.module.scss';
import moment from 'moment';
import DOMPurify from 'dompurify';

export interface IEventDetailsPanelProps {
    isOpen: boolean;
    onDismiss: () => void;
    event: IEventDetails | null;
}

export interface IEventDetails {
    title: string;
    start: Date;
    end: Date;
    color?: string;
    description?: string;
    location?: string;
    id?: string;
    url?: string; // Optional: for 'Open in SharePoint' link
}

export const EventDetailsPanel: React.FunctionComponent<IEventDetailsPanelProps> = (props) => {
    const { isOpen, onDismiss, event } = props;

    if (!event) {
        return null;
    }

    // Sanitize description if it contains HTML
    const safeDescription = event.description ? DOMPurify.sanitize(event.description) : 'No description available.';

    // Format date range
    const formatEventDate = (start: Date, end: Date) => {
        const startDate = moment(start);
        const endDate = moment(end);
        const isSameDay = startDate.isSame(endDate, 'day');

        if (isSameDay) {
            return `${startDate.format('dddd, MMMM D, YYYY')} • ${startDate.format('h:mm A')} - ${endDate.format('h:mm A')}`;
        } else {
            return `${startDate.format('MMM D, h:mm A')} - ${endDate.format('MMM D, h:mm A')}`;
        }
    };

    const onRenderHeader = () => {
        return (
            <div className={styles.headerContainer}>
                <div className={styles.colorStrip} style={{ backgroundColor: event.color || '#0078d4' }} />
                <div className={styles.headerContent}>
                    <h2 className={styles.headerTitle} role="heading" aria-level={1}>{event.title}</h2>
                    <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        ariaLabel="Close popup"
                        onClick={onDismiss}
                        className={styles.closeButton}
                    />
                </div>
            </div>
        );
    };

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            closeButtonAriaLabel="Close"
            isLightDismiss={true}
            onRenderHeader={onRenderHeader}
            hasCloseButton={false} // We implement our own in custom header
            className={styles.eventPanel}
        >
            <div className={styles.panelBody}>
                <div className={styles.metaSection}>
                    <div className={styles.metaRow}>
                        <Icon iconName="DateTime" className={styles.metaIcon} />
                        <span className={styles.metaText}>{formatEventDate(event.start, event.end)}</span>
                    </div>
                    {event.location && (
                        <div className={styles.metaRow}>
                            <Icon iconName="MapPin" className={styles.metaIcon} />
                            <span className={styles.metaText}>{event.location}</span>
                        </div>
                    )}
                </div>

                <div className={styles.descriptionSection}>
                    <h3 className={styles.sectionTitle}>Description</h3>
                    <div
                        className={styles.descriptionContent}
                        dangerouslySetInnerHTML={{ __html: safeDescription }}
                    />
                </div>

                {event.url && (
                    <div className={styles.actionSection}>
                        <a href={event.url} target="_blank" rel="noreferrer" className={styles.actionLink}>
                            Open in SharePoint <Icon iconName="OpenInNewWindow" />
                        </a>
                    </div>
                )}
            </div>
        </Panel>
    );
};
