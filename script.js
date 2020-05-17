const togglePassword = document.getElementsByClassName('show-password')[0];
const newPass = document.getElementsByName('new')[0];

togglePassword.addEventListener('click', togglePass);


function togglePass(){
	let type = newPass.getAttribute('type');


	if(type == 'password'){
		newPass.setAttribute('type','text');
		togglePassword.children[0].className = 'fas fa-eye-slash';
	}else{
		newPass.setAttribute('type','password');
		togglePassword.children[0].className = 'fas fa-eye';
	}
}

function dim(bool)
{
    if (typeof bool=='undefined') bool=true; // so you can shorten dim(true) to dim()
    document.getElementById('dimmer').style.display=(bool?'block':'none');
}   

async function sendRequest(e){
	
	e.preventDefault();
	
	let loading = document.getElementById('loading');
	let data = $('form').serialize();
	let alert;

	// dim
	dim(true);

	const request = await fetch('http://10.43.81.3/server.asp',{
		method: 'POST',
		headers: {
			'Content-type': 'application/x-www-form-urlencoded',
			'Accept': 'application/json'
		},
		body: data
	});

	const response = await request.json();

	if('Success' in response){
		dim(false);
		alert = document.getElementsByClassName('alert-success')[0];
		alert.innerHTML = response.Success;
		alert.removeAttribute('hidden');
	}else{
		dim(false);
		alert = document.getElementsByClassName('alert-danger')[0];
		alert.innerHTML = response.Error;
		alert.removeAttribute('hidden');
	}
	window.scrollTo(0,0);

	setTimeout(function(){
	    alert.setAttribute('hidden',true);
	}, 8000);
}

