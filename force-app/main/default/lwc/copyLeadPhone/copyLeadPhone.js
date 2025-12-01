import { LightningElement, api, wire } from 'lwc';
import { getRecord, getFieldValue } from 'lightning/uiRecordApi';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import PHONE_FIELD from '@salesforce/schema/Lead.Phone';

export default class CopyLeadPhone extends LightningElement {
    @api recordId;

    @wire(getRecord, { recordId: '$recordId', fields: [PHONE_FIELD] })
    lead;

    get phoneNumber() {
        return getFieldValue(this.lead.data, PHONE_FIELD);
    }

    handleCopy() {
        if (!this.lead || this.lead.error) {
            this.showToast('שגיאה בטעינת הנתונים', 'לא ניתן היה לטעון את מספר הטלפון.', 'error');
            return;
        }

        const phone = this.phoneNumber;

        if (!phone) {
            this.showToast('אין מספר טלפון', 'לא נמצא מספר טלפון ברשומת הליד.', 'warning');
            return;
        }

        this.copyToClipboard(phone)
            .then(() => {
                this.showToast('הועתק בהצלחה', `המספר ${phone} הועתק ללוח.`, 'success');
            })
            .catch(() => {
                this.showToast('שגיאה בהעתקה', 'לא הייתה אפשרות להעתיק את המספר.', 'error');
            });
    }

    copyToClipboard(text) {
        if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
            return navigator.clipboard.writeText(text);
        }

        return new Promise((resolve, reject) => {
            // Fallback using a temporary textarea element
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.setAttribute('readonly', '');
            textarea.style.position = 'absolute';
            textarea.style.left = '-9999px';
            document.body.appendChild(textarea);

            const selection = document.getSelection();
            const selectedRange = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;

            textarea.select();
            const success = document.execCommand('copy');
            document.body.removeChild(textarea);

            if (selectedRange) {
                selection.removeAllRanges();
                selection.addRange(selectedRange);
            }

            if (success) {
                resolve();
            } else {
                reject();
            }
        });
    }

    showToast(title, message, variant) {
        this.dispatchEvent(
            new ShowToastEvent({
                title,
                message,
                variant
            })
        );
    }
}
