import React, {useState} from "react";

const MultipleSelect = () => {
    const [selected, setSelected] = useState([]);

    React.useEffect(() => {
        Office.onReady(function(){
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.SelectedItemsChanged,
                handleSelectedItemChanged,
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.log('Failed to add item changed handler:', result.error);
                    } else {
                        console.log('Item changed handler added successfully');
                    }
                }
            );
        });
    }, []);

    function handleSelectedItemChanged() {
        Office.context.mailbox.getSelectedItemsAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error(asyncResult.error);
              return;
            }
            let temp = [];
            asyncResult.value.forEach((item) => {
                temp.push(item);
            });
            setSelected(temp);
          });
    }

    return (
        <div>
            <h2>Selected Items</h2>
            <ul>
                {selected.map((item) => {
                    return <li>{item.subject}</li>
                })}
            </ul>
        </div>
    );
};

export default MultipleSelect;