defmodule SymphonyElixir.RunCapabilities do
  @moduledoc """
  Builds the runtime capability summary sent at the start of an agent run.
  """

  alias SymphonyElixir.Config

  @type status :: :available | :unavailable | :unknown
  @type capability :: %{
          status: status(),
          detail: String.t()
        }
  @type t :: %{
          git_writes: capability(),
          authenticated_browser_cdp: capability(),
          smooth_task_login_profile: capability()
        }

  @spec summary(Path.t(), keyword()) :: t()
  def summary(workspace, opts \\ []) when is_binary(workspace) do
    settings = Keyword.get_lazy(opts, :settings, &Config.settings!/0)

    %{
      git_writes: git_write_capability(workspace, Keyword.get(opts, :worker_host)),
      authenticated_browser_cdp:
        declared_capability(
          settings.run_capabilities.authenticated_browser_cdp,
          "not declared by workflow config or SYMPHONY_AUTHENTICATED_BROWSER_CDP"
        ),
      smooth_task_login_profile:
        declared_capability(
          settings.run_capabilities.smooth_task_login_profile,
          "not declared by workflow config or SYMPHONY_SMOOTH_TASK_LOGIN_PROFILE"
        )
    }
  end

  @spec to_prompt(t()) :: String.t()
  def to_prompt(summary) when is_map(summary) do
    """
    Run capability summary:

    - Local `.git` writes: #{format_capability(summary.git_writes)}
    - Authenticated browser/CDP session: #{format_capability(summary.authenticated_browser_cdp)}
    - `smooth_task` Oneleet/Sprinto login profile: #{format_capability(summary.smooth_task_login_profile)}

    Use this summary when deciding whether to continue normally, use a fallback, or park the ticket for human help.
    """
    |> String.trim_trailing()
  end

  defp git_write_capability(_workspace, worker_host) when is_binary(worker_host) do
    unknown("remote worker selected; local `.git` write probe skipped")
  end

  defp git_write_capability(workspace, _worker_host) do
    with {:ok, git_dir} <- git_dir_for_workspace(workspace),
         :ok <- ensure_git_dir(git_dir),
         :ok <- ensure_no_existing_index_lock(git_dir) do
      probe_index_lock(Path.join(git_dir, "index.lock"))
    else
      {:error, detail} -> unavailable(detail)
    end
  end

  defp git_dir_for_workspace(workspace) do
    dot_git = Path.join(workspace, ".git")

    cond do
      File.dir?(dot_git) ->
        {:ok, dot_git}

      File.regular?(dot_git) ->
        dot_git_file_to_dir(dot_git, workspace)

      true ->
        {:error, "no `.git` directory or gitdir file found"}
    end
  end

  defp dot_git_file_to_dir(dot_git, workspace) do
    case File.read(dot_git) do
      {:ok, contents} ->
        contents
        |> String.split("\n", parts: 2)
        |> List.first()
        |> parse_gitdir_line(workspace)

      {:error, reason} ->
        {:error, "could not read `.git` file: #{format_reason(reason)}"}
    end
  end

  defp parse_gitdir_line("gitdir:" <> gitdir, workspace) do
    {:ok, Path.expand(String.trim(gitdir), workspace)}
  end

  defp parse_gitdir_line(_line, _workspace), do: {:error, "`.git` file did not contain a gitdir pointer"}

  defp ensure_git_dir(git_dir) do
    if File.dir?(git_dir) do
      :ok
    else
      {:error, "resolved gitdir is not a directory"}
    end
  end

  defp ensure_no_existing_index_lock(git_dir) do
    if File.exists?(Path.join(git_dir, "index.lock")) do
      {:error, "`.git/index.lock` already exists"}
    else
      :ok
    end
  end

  defp probe_index_lock(index_lock) do
    case File.write(index_lock, "symphony git write capability probe\n", [:exclusive]) do
      :ok ->
        case File.rm(index_lock) do
          :ok -> available("created and removed `.git/index.lock` probe")
          {:error, reason} -> unavailable("created `.git/index.lock` probe but cleanup failed: #{format_reason(reason)}")
        end

      {:error, reason} ->
        unavailable("could not create `.git/index.lock` probe: #{format_reason(reason)}")
    end
  end

  defp declared_capability(nil, unknown_detail), do: unknown(unknown_detail)
  defp declared_capability("", unknown_detail), do: unknown(unknown_detail)

  defp declared_capability(true, _unknown_detail),
    do: available("declared by workflow config or environment")

  defp declared_capability(false, _unknown_detail),
    do: unavailable("declared by workflow config or environment")

  defp declared_capability(value, unknown_detail) when is_binary(value) do
    case normalize_declared_status(value) do
      :available -> available("declared by workflow config or environment")
      :unavailable -> unavailable("declared by workflow config or environment")
      :unknown -> unknown("declared as unknown by workflow config or environment")
      :missing -> unknown(unknown_detail)
    end
  end

  defp normalize_declared_status(value) when is_binary(value) do
    case value |> String.trim() |> String.downcase() do
      "" -> :missing
      value when value in ["1", "true", "yes", "y", "available", "enabled", "expected"] -> :available
      value when value in ["0", "false", "no", "n", "unavailable", "disabled", "missing"] -> :unavailable
      value when value in ["unknown", "not_declared", "not-declared", "unset"] -> :unknown
      _other -> :available
    end
  end

  defp format_capability(%{status: status, detail: detail}) do
    "#{status_to_text(status)} (#{detail})"
  end

  defp status_to_text(:available), do: "available"
  defp status_to_text(:unavailable), do: "unavailable"
  defp status_to_text(:unknown), do: "unknown"

  defp available(detail), do: %{status: :available, detail: detail}
  defp unavailable(detail), do: %{status: :unavailable, detail: detail}
  defp unknown(detail), do: %{status: :unknown, detail: detail}

  defp format_reason(reason), do: Atom.to_string(reason)
end
