function deleteLastElement()
{

       /* ��ȡ��ԱԪ���б� */
       var employeeList = document.getElementsByTagName('employee');
       if (employeeList.length >> 0)
         { // �������һ����Ա������ɾ��
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
       /* ������ԱԪ��*/
       var newEmployee = document.createElement('employee');
       /* ������Ԫ�ؼ����ı���һ��������ƴ�� */
       var newName = document.createElement('name');
       var newNameText = document.createTextNode(name);
   //����ı�������
       newName.appendChild(newNameText);
       newEmployee.appendChild(newName);
       var newTitle = document.createElement('title');
       var newTitleText = document.createTextNode(title);
    //��ӱ����ı�������
       newTitle.appendChild(newTitleText);
       newEmployee.appendChild(newTitle);
       var newPhone = document.createElement('phone');
       var newPhoneText = document.createTextNode(phone);
    //��ӵ绰������
       newPhone.appendChild(newPhoneText);
       newEmployee.appendChild(newPhone);
       var newEmail = document.createElement('email');
       var newEmailText = document.createTextNode(email);
//���E-mail������
       newEmail.appendChild(newEmailText);
       newEmployee.appendChild(newEmail);
       /* ���ĵ���׷��ȫ����¼ */
       rootElement.appendChild(newEmployee);
}
