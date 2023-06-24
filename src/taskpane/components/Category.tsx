import React, {useState} from "react";

const Category = (props) => {
    const [category, setCategory] = useState(props.category);
    
    function getCategories() {
        Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const categories = asyncResult.value;
            if (categories && categories.length > 0) {
              console.log("Categories assigned to this item:");
              console.log(JSON.stringify(categories));
              document.getElementById("current_category").innerHTML = categories.map((category) => category.displayName).join(", ");
            } else {
              console.log("There are no categories assigned to this item.");
            }
          } else {
            console.error(asyncResult.error);
          }
        });
      }
      
      function addCategories() {
        // Note: In order for you to successfully add a category,
        // it must be in the mailbox categories master list.
      
        Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const masterCategories = asyncResult.value;
            if (masterCategories && masterCategories.length > 0) {
              // Grab the first category from the master list.
              const categoryToAdd = [masterCategories[0].displayName];
              Office.context.mailbox.item.categories.addAsync(categoryToAdd, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log(`Successfully assigned category '${categoryToAdd}' to item.`);
                } else {
                  console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
                }
              });
            } else {
              console.log(
                "There are no categories in the master list on this mailbox. You can add categories using Office.context.mailbox.masterCategories.addAsync."
              );
            }
          } else {
            console.error(asyncResult.error);
          }
        });
      }
      
      function removeCategories() {
        Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const categories = asyncResult.value;
            if (categories && categories.length > 0) {
              // Grab the first category assigned to this item.
              const categoryToRemove = [categories[0].displayName];
              Office.context.mailbox.item.categories.removeAsync(categoryToRemove, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log(`Successfully unassigned category '${categoryToRemove}' from this item.`);
                } else {
                  console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
                }
              });
            } else {
              console.log("There are no categories assigned to this item.");
            }
          } else {
            console.error(asyncResult.error);
          }
        });
      }

    return (
        <div>
            <h2>Category</h2>
            <div>Current category:
                <div id="current_category"></div>
            </div>
            <button onClick={getCategories}>Get Categories</button>
            <button onClick={addCategories}>Add Categories</button>
            <button onClick={removeCategories}>Remove Categories</button>
        </div>
    );
};

export default Category;