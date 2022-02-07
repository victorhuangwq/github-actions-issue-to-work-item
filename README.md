# Sync GitHub issue to Azure DevOps work item

Create a work item in Azure DevOps when a GitHub issue gets a particular label.

## Inputs

### `label`

The label that needs to be added to issues for ADO work items to be created.

### `ado_organization`

The name of the ADO organization where work items are to be created.

### `ado_project`

The name of the ADO project within the organization.

### `ado_tags`

Optional tags to be added to the work item (separated by semi-colons).

### `parent_work_item`

Optional work item number to parent the newly created work item.

### `ado_area_path`

An area path to put the work item under.

## Outputs

### `id`

The id of the Work Item created or updated

## Environment variables

The following environment variables need to be provided to the action:

* `ado_token`: an [Azure Personal Access Token](https://docs.microsoft.com/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate) with "read & write" permission for Work Item.
* `github_token`: a GitHub Personal Access Token with "repo" permissions.

## Example usage

```yaml
name: Sync issue to Azure DevOps work item

on:
  issues:
    types:
      [labeled]

jobs:
  sync:
    runs-on: ubuntu-latest
    steps:
      - uses: captainbrosset/github-actions-issue-to-work-item@patrick
        env:
          ado_token: "${{ secrets.ADO_PERSONAL_ACCESS_TOKEN }}"
          github_token: "${{ secrets.GH_PERSONAL_ACCESS_TOKEN }}"
        with:
          label: "tracked"
          ado_organization: "ado_organization_name"
          ado_project: "your_project_name"
          ado_tags: "githubSync"
          parent_work_item: 123456789
          ado_area_path: "optional_area_path"
```
