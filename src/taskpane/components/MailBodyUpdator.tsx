import React, {useState} from "react";

const MailBodyUpdator = (props) => {
    const [MailBody, setMailBody] = useState(props.MailBody);

    function prependMailBody() {
        Office.context.mailbox.item.body.setSelectedDataAsync && Office.context.mailbox.item.body.setSelectedDataAsync(
            MailBody,
            {coercionType: Office.CoercionType.Html},
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Data set successfully");
                } else {
                    console.error("Failed to set data", asyncResult.error);
                }
            }
        );
        // Office.context.mailbox.item.subject.setAsync("Hello world!", function (asyncResult) {
        //     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        //         console.log("Subject set successfully");
        //     } else {
        //         console.error("Failed to set subject", asyncResult.error);
        //     }
        // });
    }

    function isSetSupported(minVersion: string) {
        return Office.context.requirements.isSetSupported("Mailbox", minVersion)
    }

    return (
        <div>
            <div>Mail Body Manipulator</div>
            <button onClick={prependMailBody}>Replace selected(web-only)</button>
        </div>
    ); 
};

export default MailBodyUpdator;