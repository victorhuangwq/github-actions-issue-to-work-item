const core = require(`@actions/core`);
const github = require(`@actions/github`);
const azdev = require(`azure-devops-node-api`);

async function main() {
	const payload = github.context.payload;

	if (payload.action !== 'labeled' || payload.label.name !== core.getInput('label')) {
		console.log(`Either action was not 'labeled'or label was not ${core.getInput('label')}. Nothing to do.`);
		return;
	}

	let adoClient = null;

	try {
		const orgUrl = "https://dev.azure.com/" + core.getInput('ado_organization');
		const adoAuthHandler = azdev.getPersonalAccessTokenHandler(process.env.ado_token);
		const adoConnection = new azdev.WebApi(orgUrl, adoAuthHandler);
		adoClient = await adoConnection.getWorkItemTrackingApi();
	} catch (e) {
		console.error(e);
		core.setFailed('Could not connect to ADO');
		return;
	}

	try {
		// go check to see if work item already exists in azure devops or not
		// based on the title and tags.
		console.log("Check to see if work item already exists");
		let adoId = await findAdoId(payload.issue.number, adoClient);
		if (adoId === null) {
			console.log("Could not find existing ADO workitem, creating one now");
		} else {
			console.log("Found existing ADO workitem: " + adoId + ". No need to create a new one");
			return;
		}

		// if workItem == -1 then we have an error during find
		if (adoId === -1) {
			core.setFailed("Error while finding the ADO work item");
			return;
		}

		let workItem = await create(payload, adoClient);

		// set output message
		if (workItem != null || workItem != undefined) {
			console.log(`Work item successfully created or found: ${workItem.id}`);
			core.setOutput(`id`, `${workItem.id}`);
		}
	} catch (error) {
		console.log("Error: " + error);
		core.setFailed();
	}
}

function formatTitle(githubIssue) {
	return "[GitHub #" + githubIssue.number + "] " + githubIssue.title;
}

async function formatDescription(githubIssue) {
	console.log('Creating a description based on the github issue');
	const octokit = new github.GitHub(process.env.github_token);
	const bodyWithMarkdown = await octokit.markdown.render({ text: githubIssue.body });

	return '________________________________________________________<br>' +
		'<em>This item was auto-opened from GitHub <a href="' +
		githubIssue.html_url +
		'" target="_new">issue #' +
		githubIssue.number +
		"</a></em><br>" +
		"It won't auto-update when the GitHub issue changes so please check the issue for updates.<br><br>" +
		"<strong>Initial description from GitHub (check issue for more info):</strong><br><br>" +
		bodyWithMarkdown.data;
}

async function create(payload, adoClient) {
	const botMessage = await formatDescription(payload.issue);
	const shortRepoName = payload.repository.full_name.split("/")[1];
	const tags = core.getInput("ado_tags") ? core.getInput("ado_tags") + ";" + shortRepoName : shortRepoName;
	const isFeature = payload.issue.labels.some((label) => label === 'enhancement' || label === 'feature' || label == 'feature request');

	console.log(`Starting to create work item for GitHub issue #${payload.issue.number}`);

	const patchDocument = [
		{
			op: "add",
			path: "/fields/System.Title",
			value: formatTitle(payload.issue),
		},
		{
			op: "add",
			path: "/fields/System.Description",
			value: botMessage,
		},
		{
			op: "add",
			path: "/fields/Microsoft.VSTS.TCM.ReproSteps",
			value: botMessage,
		},
		{
			op: "add",
			path: "/fields/System.Tags",
			value: tags,
		},
		{
			op: "add",
			path: "/relations/-",
			value: {
				rel: "Hyperlink",
				url: payload.issue.html_url,
			},
		}
	];

	if (core.getInput('parent_work_item')) {
		let parentUrl = "https://dev.azure.com/" + core.getInput('ado_organization');
		parentUrl += '/_workitems/edit/' + core.getInput('parent_work_item');

		patchDocument.push({
			op: "add",
			path: "/relations/-",
			value: {
				rel: "System.LinkTypes.Hierarchy-Reverse",
				url: parentUrl,
				attributes: {
					comment: ""
				}
			}
		});
	}

	patchDocument.push({
		op: "add",
		path: "/fields/System.AreaPath",
		value: core.getInput('ado_area_path'),
	});

	let workItemSaveResult = null;

	try {
		console.log('Creating work item');
		workItemSaveResult = await adoClient.createWorkItem(
			(customHeaders = []),
			(document = patchDocument),
			(project = core.getInput('ado_project')),
			(type = isFeature ? 'Scenario' : 'Bug'),
			(validateOnly = false),
			(bypassRules = false)
		);

		// if result is null, save did not complete correctly
		if (workItemSaveResult == null) {
			workItemSaveResult = -1;

			console.log("Error: createWorkItem failed");
			console.log(`WIT may not be correct: ${wit}`);
			core.setFailed();
		} else {
			console.log("Work item successfully created");
		}
	} catch (error) {
		workItemSaveResult = -1;

		console.log("Error: createWorkItem failed");
		console.log(patchDocument);
		console.log(error);
		core.setFailed(error);
	}

	if (workItemSaveResult != -1) {
		console.log(workItemSaveResult);
	}

	return workItemSaveResult;
}

async function findAdoId(ghIssueId, adoClient) {
	console.log('Connecting to Azure DevOps to find work item for issue #' + ghIssueId);

	const wiql = {
		query:
			`SELECT [System.Id], [System.WorkItemType], [System.Description], [System.Title], [System.AssignedTo], [System.State], [System.Tags]
			FROM workitems 
			WHERE [System.TeamProject] = @project AND [System.Title] CONTAINS 'GitHub #' AND [System.Title] CONTAINS '${ghIssueId}' AND [System.AreaPath] = '${core.getInput('ado_area_path')}'`
	};
	console.log("ADO query: " + wiql.query);

	let queryResult = null;
	try {
		queryResult = await adoClient.queryByWiql(wiql, { project: core.getInput('ado_project') });

		// if query results = null then i think we have issue with the project name
		if (queryResult == null) {
			console.log("Error: Project name appears to be invalid");
			core.setFailed("Error: Project name appears to be invalid");
			return -1;
		}
	} catch (error) {
		console.log("Error: queryByWiql failure");
		console.log(error);
		core.setFailed(error);
		return -1;
	}
	console.log(queryResult);

	console.log("Use the first item found");
	const workItem = queryResult.workItems.length > 0 ? queryResult.workItems[0] : null;

	if (workItem != null) {
		try {
			console.log("Workitem data retrieved: " + workItem.id);
			return workItem.id;
		} catch (error) {
			console.log("Error: getWorkItem failure");
			core.setFailed(error);
			return -1;
		}
	} else {
		return null;
	}
}

main();
