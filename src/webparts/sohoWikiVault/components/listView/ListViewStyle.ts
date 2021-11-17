import { IDialogContentStyles, IDialogStyles, IButtonStyles } from 'office-ui-fabric-react';

export class ListViewStyle {
    constructor() { }
    public errorHeaderStyle: IDialogContentStyles = {
        content: {},
        inner: {},
        innerContent: {},
        title: { fontSize: 25, fontWeight: 500, padding: 20 },
        subText: { color: '#da0404', fontSize: 15, fontWeight: 400 },
        header: { backgroundColor: '#da0404' },
        button: { backgroundColor: '#da0404' },
        topButton: {}
    };
    public editDialogStyle: IDialogStyles = {
        main: {
            selectors: {
                ['@media (min-width: 480px)']: {
                    width: '65%',
                    maxWidth: '100%'
                }
            }
        },
        root: {}
    };
    public viewDialogStyle: IDialogStyles = {
        main: {
            selectors: {
                ['@media (min-width: 480px)']: {
                    width: '65%',
                    maxWidth: '65%'
                }
            }
        },
        root: {}
    };
    public viewHeaderStyle: IDialogContentStyles = {
        content: {},
        inner: {},
        innerContent: {},
        title: { fontSize: 20, fontWeight: 500, padding: 15 },
        subText: { fontSize: 15, fontWeight: 400 },
        header: {},
        button: {},
        topButton: {}
    };
    public editDialogControlStyle: IDialogContentStyles = {
        content: {},
        inner: {},
        innerContent: {},
        title: { fontSize: '20px', fontWeight: "500", padding: 10 },
        subText: {},
        header: {},
        button: {},
        topButton: {}
    };
    public deleteButtonStyle: IButtonStyles = {
        root: { backgroundColor: '#da0404' },
        rootCheckedHovered: { backgroundColor: '#da0404' },
        rootHovered: { backgroundColor: '#da0404' },
        rootChecked: { backgroundColor: '#da0404' },
    };
}