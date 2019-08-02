const request = require('request');
const date_format = require('dateformat');

const office_365_tenant = 'my-company.onmicrosoft.com';
const sharepoint_site_for_faxes = 'Site01';
const sharepoint_document_library_for_faxes = 'Incoming Faxes';

const graph_authentication_endpoint = 'https://login.microsoftonline.com/' + office_365_tenant + '/oauth2/v2.0/token';

const request_options = {
	grant_type: 'client_credentials',
	client_id: process.env.MS_GRAPH_APP_ID,
	client_secret: process.env.MS_GRAPH_APP_SECRET,
	scope: 'https://graph.microsoft.com/.default'
};

// Just some sample data here
var fax_filename = 'test_file.pdf';
var media_url = 'https://file-examples.com/wp-content/uploads/2017/10/file-sample_150kB.pdf';
var phone_number = '+10123456789'

const FAX_NUMBERS = {
	'+10123456789': {
		description: 'first fax number',
		sharepoint_folder: 'faxfolder1'
	},
	'+11234567890': {
		description: 'second fax number',
		sharepoint_folder: 'faxfolder2'
	},
	'+19876543210': {
		description: 'third fax number',
		sharepoint_folder: 'faxfolder1'
	}
};

const ms_graph_connect = () => {
	return new Promise(function(resolve, reject) {
		request.post({ url: graph_authentication_endpoint, form: request_options }, function(err, response, body) {
			if (err) {
				return reject(err);
			} else {
				try {
					let parsed_body = JSON.parse(body);
					if (parsed_body.error_description) {
						reject('Error=' + parsed_body.error_description);
					} else {
						resolve({ token: parsed_body.access_token });
					}
				} catch (e) {
					reject(e);
				}
			}
		});
	});
};

const get_sharepoint_site_id = (data) => {
	return new Promise(function(resolve, reject) {
		request.get(
			{
				url:
					'https://graph.microsoft.com/v1.0/sites/' +
					office_365_tenant.split('.')[0] +
					'.sharepoint.com:/sites/' +
					sharepoint_site_for_faxes,
				headers: {
					Authorization: 'Bearer ' + data.token
				}
			},
			function(err, response, body) {
				if (err) {
					reject(err);
				} else if (response.statusCode !== 200) {
					reject(body);
				} else {
					try {
						let parsed_body = JSON.parse(body);
						data.site_id = parsed_body.id;
						resolve(data);
					} catch (e) {
						reject(e);
					}
				}
			}
		);
	});
};

const get_sharepoint_drive_id = (data) => {
	return new Promise(function(resolve, reject) {
		request.get(
			{
				url: 'https://graph.microsoft.com/v1.0/sites/' + data.site_id + '/drives',
				headers: {
					Authorization: 'Bearer ' + data.token
				}
			},
			function(err, response, body) {
				if (err) {
					reject(err);
				} else if (response.statusCode !== 200) {
					reject(body);
				} else {
					try {
						let parsed_body = JSON.parse(body);
						parsed_body.value.forEach(function(drive) {
							if (
								drive.name === sharepoint_document_library_for_faxes &&
								drive.driveType === 'documentLibrary'
							) {
								data.drive_id = drive.id;
								resolve(data);
							}
						});
						reject({
							error: 'Document library "' + sharepoint_document_library_for_faxes + '" not found.'
						});
					} catch (e) {
						reject(e);
					}
				}
			}
		);
	});
};

const get_folder_id = (data) => {
	return new Promise(function(resolve, reject) {
		request.get(
			{
				url: 'https://graph.microsoft.com/v1.0/drives/' + data.drive_id + '/root' + '/children',
				headers: {
					Authorization: 'Bearer ' + data.token
				}
			},
			function(err, response, body) {
				if (err) {
					reject(err);
				} else if (response.statusCode !== 200) {
					reject(body);
				} else {
					try {
						let parsed_body = JSON.parse(body);
						parsed_body.value.forEach(function(folder) {
							if (folder.name === FAX_NUMBERS[phone_number].sharepoint_folder) {
								data.folder_id = folder.id;
								resolve(data);
							}
						});
						reject({
							error:
								'Folder "' +
								FAX_NUMBERS[phone_number].sharepoint_folder +
								'" is not found in "' +
								sharepoint_document_library_for_faxes +
								'" document library.'
						});
					} catch (e) {
						reject(e);
					}
				}
			}
		);
	});
};

const upload_file_to_sharepoint = (data) => {
	return new Promise(function(resolve, reject) {
		request
			.get({ url: media_url }, function(err, response, body) {
				if (err) {
					reject(err);
				} else if (response.statusCode !== 200) {
					reject(body);
				}
			})
			.pipe(
				request.put(
					{
						url:
							'https://graph.microsoft.com/v1.0/drives/' +
							data.drive_id +
							'/items/' +
							data.folder_id +
							':/' +
							fax_filename +
							':/content',
						headers: {
							Authorization: 'Bearer ' + data.token,
							'Content-Type': 'text/plain'
						}
					},
					function(err, response, body) {
						if (err) {
							reject(err);
						} else if (response.statusCode !== 201) {
							reject(body);
						} else {
							try {
								parsed_body = JSON.parse(body);
								resolve(parsed_body.webUrl);
							} catch (e) {
								reject(e);
							}
						}
					}
				)
			);
	});
};

exports.handler = function(context, event, callback) {
    media_url = event.MediaUrl;
    fax_filename = date_format(new Date(), "UTC:yyyy-mm-dd-HH-MM-ss") + '-from-' + event.From.substring(2) + '-to-' + event.To.substring(2) + '.pdf';
    phone_number = event.To;
    
    if (!FAX_NUMBERS.hasOwnProperty(phone_number)) {
        callback('Phone number ' + phone_number + ' is not configured in FAX_NUMBERS constant');
    }

	  ms_graph_connect()
	      .then(get_sharepoint_site_id)
	      .then(get_sharepoint_drive_id)
	      .then(get_folder_id)
	      .then(upload_file_to_sharepoint)
	      .then(function(data) {
  		      callback(null, data);
	      })
	      .catch(function(err) {
		        callback(err);
	      });
};
