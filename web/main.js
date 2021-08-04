
async function send() {

    var letter_title = document.querySelector("#letter_title").value;
    var html_path = document.querySelector("#html_path").value;
    var csv_path = document.querySelector("#csv_path").value;
    var sender_mail = document.querySelector("#sender_mail").value;
    var sender_password = document.querySelector("#sender_password").value;
    var mail_csv_order = document.querySelector("#mail_csv_order").value;
    var name_csv_order = document.querySelector("#name_csv_order").value;
    var mail_content1_order = document.querySelector("#mail_content1_order").value;
    var mail_content2_order = document.querySelector("#mail_content2_order").value;
    var attached_dir_path = document.querySelector("#attached_dir_path").value;
    var attached_csv_order = document.querySelector("#attached_csv_order").value;

    document.querySelector('#result').textContent = "寄信中..."
    result = await eel.main(letter_title,html_path,csv_path,sender_mail,sender_password,mail_csv_order,name_csv_order,mail_content1_order,mail_content2_order,attached_dir_path,attached_csv_order)();
    document.querySelector('#result').textContent = result;

}

async function send_authorization() {
    var acad_authorization = document.querySelector("#authorization").value;
    result = await eel.comfirm_authorization(acad_authorization)();
    if (result=="success"){
    	confirm("授權成功");
        document.location.href="main.html"
    }else if(result=="failure"){
        document.querySelector('#authorization_result').textContent = "授權失敗";
    }else{
    	document.querySelector('#authorization_result').textContent = result;
	}
}

// 如果檔案中有寫入授權碼，就直接進到授權頁面，不用再輸入一次授權碼
async function check_local_ket(){
    result = await eel.have_use_authorization()();
    if (result=="success"){
        document.location.href="main.html"
    }else{
    }
}

