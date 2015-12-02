function deleteLastElement()
{

       /* 获取雇员元素列表 */
       var employeeList = document.getElementsByTagName('employee');
       if (employeeList.length >> 0)
         { // 查找最后一个雇员并将其删除
          var toDelete = employeeList.item(employeeList.length-1);
          document.documentElement.removeChild(toDelete);
         }
       else
         alert('No employee elements to delete');
}
function addElement()
{
       var rootElement = document.documentElement;
       var name = document.getElementById('namefield').value;
       var title = document.getElementById('titlefield').value;
       var phone = document.getElementById('phonefield').value;
       var email = document.getElementById('emailfield').value;
       /* 创建雇员元素*/
       var newEmployee = document.createElement('employee');
       /* 创建子元素及其文本并一个个进行拼接 */
       var newName = document.createElement('name');
       var newNameText = document.createTextNode(name);
   //添加文本、名称
       newName.appendChild(newNameText);
       newEmployee.appendChild(newName);
       var newTitle = document.createElement('title');
       var newTitleText = document.createTextNode(title);
    //添加标题文本、名称
       newTitle.appendChild(newTitleText);
       newEmployee.appendChild(newTitle);
       var newPhone = document.createElement('phone');
       var newPhoneText = document.createTextNode(phone);
    //添加电话、名称
       newPhone.appendChild(newPhoneText);
       newEmployee.appendChild(newPhone);
       var newEmail = document.createElement('email');
       var newEmailText = document.createTextNode(email);
//添加E-mail、名称
       newEmail.appendChild(newEmailText);
       newEmployee.appendChild(newEmail);
       /* 向文档中追加全部记录 */
       rootElement.appendChild(newEmployee);
}
