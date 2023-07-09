using Nuke.Common;
using Nuke.Common.CI.GitHubActions;

[GitHubActions(
    "build",
    GitHubActionsImage.UbuntuLatest,
    AutoGenerate = true,
    PublishArtifacts = true,
    EnableGitHubToken = true,
    OnPushBranches = new[] { DevelopBranch },
    InvokedTargets = new[] { nameof(Compile) },
    CacheKeyFiles = new string[0]
)]
[GitHubActions(
    "deploy",
    GitHubActionsImage.UbuntuLatest,
    AutoGenerate = true,
    PublishArtifacts = true,
    EnableGitHubToken = true,
    OnPushBranches = new[] { MainBranch,PreviewBranch },
    InvokedTargets = new[] { nameof(Deploy) },
    ImportSecrets = new[] { nameof(NuGetApiKey), nameof(MyGetApiKey) },
    CacheKeyFiles = new string[0]
)]
partial class Build
{
    GitHubActions GitHubActions => GitHubActions.Instance;
    
    // private Target Release => _ => _
    //     .Description("Release")
    //     .Executes(() =>
    //     {
    //         GitReleaseManagerCreate(new Nuke.Common.Tools.GitReleaseManager.GitReleaseManagerCreateSettings());
    //         //GitReleaseManagerAddAssets( )
    //     });
}