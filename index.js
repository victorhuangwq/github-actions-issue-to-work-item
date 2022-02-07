const core = require(`@actions/core`);
const github = require(`@actions/github`);
const azdev = require(`azure-devops-node-api`);

async function main() {
	const payload = github.context.payload;

	if (payload.action !== 'labeled' || !core.getInput('label')) {
		core.setFailed(`Action not supported: ${payload.action}. Only 'labeled' is supported.`);
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
		let workItem = await find(payload.issue.number, adoClient);
		if (workItem === null) {
			console.log("Could not find existing ADO workitem");
		} else {
			console.log("Found existing ADO workitem: " + workItem.id);
			return;
		}

		// if workItem == -1 then we have an error during find
		if (workItem === -1) {
			core.setFailed("Error while finding the ADO work item");
			return;
		}

		if (payload.label.name === core.getInput('label')) {
			workItem = await create(payload, adoClient);
		}

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

	return '<em>This item was auto-opened from GitHub <a href="' +
		githubIssue.html_url +
		'" target="_new">issue #' +
		githubIssue.number +
		"</a></em><br>" +
		"It won't auto-update when the GitHub issue changes so please check the issue for updates.<br><br>" +
		"<strong>Description from GitHub:</strong><br><br>" +
		bodyWithMarkdown.data;
}

async function create(payload, adoClient) {
	const botMessage = await formatDescription(payload.issue);
	const shortRepoName = payload.repository.full_name.split("/")[1];
	const tags = core.getInput("ado_tags") ? core.getInput("ado_tags") + ";" + shortRepoName : shortRepoName;
	const isFeature = payload.issue.labels.some((label) => label === 'enhancement');

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
		patchDocument.push({
			op: "add",
			path: "/relations/-",
			value: {
				rel: "System.LinkTypes.Hierarchy-Reverse",
				url: core.getInput('parent_work_item'),
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

async function find(ghIssueNb, adoClient) {
	console.log('Connecting to Azure DevOps to find work item for issue #' + ghIssueNb);

	const wiql = {
		query:
			`SELECT [System.Id], [System.WorkItemType], [System.Description], [System.Title], [System.AssignedTo], [System.State], [System.Tags]
			FROM workitems 
			WHERE [System.TeamProject] = @project AND [System.Title] CONTAINS '[GitHub #${ghIssueNb}]' AND [System.AreaPath] = '${core.getInput('ado_area_path')}'`
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

	console.log("Use the first item found");
	const workItem = queryResult.workItems.length > 0 ? queryResult.workItems[0] : null;

	if (workItem != null) {
		try {
			var result = await client.getWorkItem(workItem.id, null, null, 4);
			console.log("Workitem data retrieved: " + workItem.id);
			return result;
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
