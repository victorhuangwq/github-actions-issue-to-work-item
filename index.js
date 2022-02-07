const core = require(`@actions/core`);
const github = require(`@actions/github`);
const azdev = require(`azure-devops-node-api`);
const fetch = require(`node-fetch`);

async function main() {
	const payload = github.context.payload;

	// Action only runs on a label event.
	if (payload.action !== 'labeled' || !core.getInput('label')) {
		core.setFailed(`Action not supported: ${payload.action}. Only 'labeled' is supported.`);
		return;
	}

	try {
		// go check to see if work item already exists in azure devops or not
		// based on the title and tags.
		console.log("Check to see if work item already exists");
		let workItem = await find(payload.issue.number);
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
			workItem = await create(payload);
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

async function formatDescription(githubIssue, githubRepository) {
	const octokit = new github.GitHub(process.env.github_token);
	const bodyWithMarkdown = await octokit.markdown.render({ text: githubIssue.body })

	return 'This item was auto-opened from GitHub <a href="' +
		githubIssue.html_url +
		'" target="_new">issue #' +
		githubIssue.number +
		'</a> created in the <a href="' +
		githubRepository.html_url +
		'" target="_new">' +
		githubRepository.name +
		"</a>  project</br></br><b>Description from GitHub: </b></br>" +
		bodyWithMarkdown.data;
}

async function formatHistory(githubIssue, githubRepository) {
	let history =
		'GitHub <a href="' +
		githubIssue.html_url +
		'" target="_new">issue #' +
		githubIssue.number +
		'</a> labeled as ' +
		core.getInput('label') +
		' in <a href="' +
		githubRepository.html_url +
		'" target="_new">' +
		githubRepository.full_name +
		"</a>";

	const commentsUrl = `https://api.github.com/repos/${githubRepository.full_name}/issues/${githubIssue.number}/comments`;
	const comments = await fetch(commentsUrl)
		.then((res) => res.json())
		.catch(err => console.log(err));
	for (const i in comments) {
		const comment = comments[i];
		history +=
			'</br></br>GitHub <a href="' +
			comment.html_url +
			'" target="_new">comment</a> by ' +
			comment.user.login +
			' on ' +
			comment.created_at +
			':</br>' +
			comment.body;
	}
	return history;
}

async function create(payload) {
	const botMessage = await formatDescription(payload.issue, payload.repository, env);
	const shortRepoName = payload.repository.full_name.split("/")[1];
	const tags = core.getInput("ado_tags") ? core.getInput("ado_tags") + ";" + shortRepoName : shortRepoName;
	const isFeature = payload.issue.labels.some((label) => label === 'enhancement');

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

	// Migrate issue history
	const history = await formatHistory(payload.issue, payload.repository);
	patchDocument.push({
		op: "add",
		path: "/fields/System.History",
		value: history,
	});

	let authHandler = azdev.getPersonalAccessTokenHandler(process.env.ado_token);
	let connection = new azdev.WebApi(core.getInput('ado_organization'), authHandler);
	let client = await connection.getWorkItemTrackingApi();
	let workItemSaveResult = null;

	try {
		workItemSaveResult = await client.createWorkItem(
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

			console.log("Error: creatWorkItem failed");
			console.log(`WIT may not be correct: ${wit}`);
			core.setFailed();
		} else {
			console.log("Work item successfully created");
		}
	} catch (error) {
		workItemSaveResult = -1;

		console.log("Error: creatWorkItem failed");
		console.log(patchDocument);
		console.log(error);
		core.setFailed(error);
	}

	if (workItemSaveResult != -1) {
		console.log(workItemSaveResult);
	}

	return workItemSaveResult;
}

async function find(ghIssueNb) {
	const orgUrl = "https://dev.azure.com/" + core.getInput('ado_organization');

	let authHandler = azdev.getPersonalAccessTokenHandler(process.env.ado_token);
	let connection = new azdev.WebApi(orgUrl, authHandler);
	let client = null;
	let workItem = null;
	let queryResult = null;

	console.log("Finding workitem");
	try {
		client = await connection.getWorkItemTrackingApi();
	} catch (error) {
		console.log(
			"Error: Connecting to organization. Check the spelling of the organization name and ensure your token is scoped correctly."
		);
		core.setFailed(error);
		return -1;
	}

	let teamContext = { project: core.getInput('ado_project') };

	let wiql = {
		query:
			`SELECT [System.Id], [System.WorkItemType], [System.Description], [System.Title], [System.AssignedTo], [System.State], [System.Tags]
			FROM workitems 
			WHERE [System.TeamProject] = @project AND [System.Title] CONTAINS '[GitHub #${ghIssueNb}]' AND [System.AreaPath] = '${core.getInput('ado_area_path')}'`
	};
	console.log("ADO query: " + wiql.query);

	try {
		queryResult = await client.queryByWiql(wiql, teamContext);

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
	workItem = queryResult.workItems.length > 0 ? queryResult.workItems[0] : null;

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
