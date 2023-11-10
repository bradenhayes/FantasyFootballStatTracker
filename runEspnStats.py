from espn_api.football import League
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
import json
import datetime


class TeamData:
    """
    Represents data for a team in the fantasy football league.

    :param team: The team object.
    :param lineup: The team's lineup for a specific week.
    :param index_column: The column index in a spreadsheet.
    :param projected: The projected score for the team.
    :param score: The actual score for the team.
    :returns: A TeamData object.
    :rtype: TeamData
    """

    def __init__(self, team, lineup, index_column, projected, score):
        self.team = team
        self.lineup = lineup
        self.index_column = index_column
        self.projected = projected
        self.score = score
        self.additional_data = {}

    def add_average_score_per_player(self, position, score):
        """
        Add the average score per player for a specific position.

        :param position: The player's position (e.g., "QB", "WR", "RB").
        :param score: The average score for the position.
        :returns: None
        :rtype: None
        """
        if "average_per_overall" not in self.additional_data:
            self.additional_data[
                "average_per_overall"
            ] = []  # Initialize as an empty list
        self.additional_data["average_per_overall"].append((position, score))

    def get_average_score_per_player(self):
        """
        Get the average score per player for each position.

        :returns: List of tuples containing position and average score.
        :rtype: list
        """
        return self.additional_data.get("average_per_overall")

    def add_power_ranking(self, ranking):
        """
        Add the power ranking for the team.

        :param ranking: The power ranking value.
        :returns: None
        :rtype: None
        """
        self.additional_data["power_ranking"] = ranking

    def get_power_ranking(self):
        """
        Get the power ranking for the team.

        :returns: The power ranking value.
        :rtype: int
        """
        return self.additional_data.get("power_ranking")

    def add_percentage_per_position(self, position, score):
        """
        Add the percentage of points for a specific position.

        :param position: The player's position (e.g., "QB", "WR", "RB").
        :param score: The percentage of points for the position.
        :returns: None
        :rtype: None
        """
        if "percentage_per_position" not in self.additional_data:
            self.additional_data[
                "percentage_per_position"
            ] = []  # Initialize as an empty list
        self.additional_data["percentage_per_position"].append((position, score))

    def get_percentage_per_position(self):
        """
        Get the percentage of points for each position.

        :returns: List of tuples containing position and percentage of points.
        :rtype: list
        """
        return self.additional_data.get("percentage_per_position")


class LeagueCreator:
    """
    Provides functionality to create a League object for accessing ESPN Fantasy Football data.

    :param league_id: The ESPN Fantasy Football league ID.
    :param year: The year of the league.
    :param espn_s2: The ESPN S2 token.
    :param swid: The SWID (Secure Web Login ID) token.
    :returns: A LeagueCreator object.
    :rtype: LeagueCreator
    """

    def __init__(self, league_id, year, espn_s2, swid):
        self.league_id = league_id
        self.year = year
        self.espn_s2 = espn_s2
        self.swid = swid

    def create_league(self):
        """
        Create a League object for the specified league and credentials.

        :returns: League object for accessing ESPN Fantasy Football data.
        :rtype: espn_api.football.League
        """
        return League(
            league_id=self.league_id,
            year=self.year,
            espn_s2=self.espn_s2,
            swid=self.swid,
        )


class FantasyFootballLeague:
    """
    Represents a fantasy football league and provides various data analysis methods.

    :param league_creator: An instance of LeagueCreator.
    :param week: The current week of the fantasy league.
    :returns: A FantasyFootballLeague object.
    :rtype: FantasyFootballLeague
    """

    def __init__(self, league_creator, week):
        self.league = league_creator
        self.week = week

    def get_teams_data(self, week_option=None):
        """
        Get data for all teams in the league for a specific week or the current week.

        :param week_option: The week number for which to fetch the data (optional).
        :returns: List of TeamData objects containing team information, lineup, scores, and additional data.
        :rtype: list
        """
        if week_option:
            box_scores = self.league.box_scores(week_option)
        else:
            box_scores = self.league.box_scores(self.week)
        teams_data = []

        for matchup in box_scores:
            for team_type in ["home", "away"]:
                team = getattr(matchup, f"{team_type}_team")
                lineup = getattr(matchup, f"{team_type}_lineup")
                projected = getattr(matchup, f"{team_type}_projected")
                score = getattr(matchup, f"{team_type}_score")
                index_column = self.league.teams.index(team) + 1

                team_data = TeamData(team, lineup, index_column, projected, score)
                teams_data.append(team_data)
        return teams_data

    def average_points_per_player(self, positions, slot_conditions=[], week=None):
        """
        Calculate the average points per player for specific positions.

        :param positions: List of positions (e.g., ["QB", "WR", "RB", "TE", "K", "D/ST"]).
        :param slot_conditions: List of slot conditions to exclude (e.g., ["BE"]).
        :param week: The week number for which to calculate the average (optional).
        :returns: List of TeamData objects with added average scores per player for specified positions.
        :rtype: list
        """
        if week:
            teams_data = self.get_teams_data(week)
            week = week
        else:
            teams_data = self.get_teams_data()
            week = self.week
        for team in teams_data:
            lineup = team.lineup
            for position in positions:
                team_player_count = 0
                team_player_stats = 0
                for player in lineup:
                    if (
                        player.position == position
                        and player.slot_position not in slot_conditions
                    ):
                        team_player_count += 1
                        if week in player.stats and "points" in player.stats[week]:
                            team_player_stats += player.stats[week]["points"]

                if team_player_count > 0:
                    average_player_points = round(
                        team_player_stats / team_player_count, 2
                    )
                    team.add_average_score_per_player(position, average_player_points)

        return teams_data

    def power_ranking_per_player(self, week=None):
        """
        Calculate the power ranking for each team in the league for a specific week or the current week.

        :param week: The week number for which to calculate the power ranking (optional).
        :returns: List of TeamData objects with added power rankings.
        :rtype: list
        """
        if week:
            power_rankings = self.league.power_rankings(week)
            teams_data = self.get_teams_data(week)
        else:
            power_rankings = self.league.power_rankings(self.week)
            teams_data = self.get_teams_data(self.week)
        for ranking in power_rankings:
            for team in teams_data:
                if ranking[1] == team.team:
                    team.add_power_ranking(ranking[0])
        return teams_data

    def percentage_of_points_by_position(
        self, positions, slot_conditions=[], week=None
    ):
        """
        Calculate the percentage of points for specific positions.

        :param positions: List of positions (e.g., ["QB", "WR", "RB", "TE", "K", "D/ST"]).
        :param slot_conditions: List of slot conditions to exclude (e.g., ["BE"]).
        :param week: The week number for which to calculate the percentage (optional).
        :returns: List of TeamData objects with added percentage of points for specified positions.
        :rtype: list
        """
        if week:
            teams_data = self.get_teams_data(week)
            week = week
        else:
            teams_data = self.get_teams_data()
            week = self.week
        for team in teams_data:
            lineup = team.lineup
            for position in positions:
                total_points_for_position = sum(
                    player.stats[week]["points"]
                    for player in lineup
                    if week in player.stats
                    and "points" in player.stats[week]
                    and player.position == position
                    and player.slot_position not in slot_conditions
                )
                # Calculate the percentage of points for this position
                total_team_points = sum(
                    player.stats[week]["points"]
                    for player in lineup
                    if week in player.stats
                    and "points" in player.stats[week]
                    and player.slot_position not in slot_conditions
                )

                if total_team_points != 0:
                    percentage_points_per_position = (
                        total_points_for_position / total_team_points
                    ) * 100
                else:
                    percentage_points_per_position = 0

                # Round the percentage to 2 decimal places
                percentage_points_per_position_rounded = round(
                    percentage_points_per_position, 2
                )
                team.add_percentage_per_position(
                    position, percentage_points_per_position_rounded
                )

        return teams_data


class GoogleSheetsManager:
    """
    Represents a Google Sheets service for interacting with Google Sheets API.

    :param credentials_path: The path to the JSON credentials file.
    :param spreadsheet_id: The ID of the Google Sheets spreadsheet.
    :returns: A GoogleSheets object.
    :rtype: GoogleSheets
    """

    def __init__(self, credentials_file, spreadsheet_id, batch_size):
        """
        Initialize the GoogleSheetsManager.

        :param credentials_file: The path to the JSON credentials file.
        :param spreadsheet_id: The ID of the Google Sheets spreadsheet.
        :param batch_size: The batch size for writing data.
        """
        self.creds = service_account.Credentials.from_service_account_file(
            credentials_file, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        self.spreadsheet_id = spreadsheet_id
        self.service = build("sheets", "v4", credentials=self.creds)
        self.batch_size = batch_size
        self.batch_data = {}  # Store batch data for each function

    def create_new_sheet(self, sheet_name):
        """
        Create a new sheet in the Google Sheets document.

        :param sheet_name: The name of the new sheet to create.
        """
        try:
            # Create a new sheet
            new_sheet_request = {
                "addSheet": {
                    "properties": {
                        "title": sheet_name,
                    }
                }
            }
            body = {"requests": [new_sheet_request]}
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id, body=body
            ).execute()

            print(f"Created a new sheet '{sheet_name}'.")
        except HttpError as e:
            print(f"Error creating a new sheet: {e}")

        # Write the accumulated team names in batches
        self.write_batch()

        print(f"Added team names to sheet '{sheet_name}'.")

    def delete_other_sheets(self, manager, sheet_to_keep_name):
        """
        Delete all sheets in the Google Sheets document except the one specified by `sheet_to_keep_name`.

        :param manager: An instance of GoogleSheetsManager.
        :param sheet_to_keep_name: The name of the sheet to keep (not deleted).
        """
        sheet_list = manager.open_spreadsheet().get("sheets", [])

        for sheet in sheet_list:
            title = sheet.get("properties", {}).get("title")
            if title != sheet_to_keep_name:
                # Create a request to delete the sheet
                delete_request = {
                    "deleteSheet": {
                        "sheetId": sheet.get("properties", {}).get("sheetId")
                    }
                }

                # Batch update to delete the sheet
                body = {"requests": [delete_request]}
                manager.service.spreadsheets().batchUpdate(
                    spreadsheetId=manager.spreadsheet_id, body=body
                ).execute()
            else:
                print(f"Keeping sheet '{sheet_to_keep_name}'.")

    def check_sheet_exists(self, sheet_name):
        """
        Check if a sheet with the given name exists in the Google Sheets document.

        :param sheet_name: The name of the sheet to check.
        :return: True if the sheet exists, False otherwise.
        :rtype: bool
        """
        try:
            spreadsheet = self.open_spreadsheet()
            if spreadsheet:
                sheets = spreadsheet.get("sheets", [])
                for sheet in sheets:
                    title = sheet.get("properties", {}).get("title")
                    if title == sheet_name:
                        return True
            return False
        except HttpError as e:
            print(f"Error checking if sheet exists: {e}")
            return False

    def open_spreadsheet(self):
        """
        Open the Google Sheets spreadsheet.

        :return: The spreadsheet data if successful, None if there is an error.
        :rtype: dict or None
        """
        try:
            spreadsheet = (
                self.service.spreadsheets()
                .get(spreadsheetId=self.spreadsheet_id)
                .execute()
            )
            return spreadsheet
        except HttpError as e:
            print(f"Error opening spreadsheet: {e}")
            return None

    def add_data_to_batch(self, sheet_name, cell_range, value, function_name):
        """
        Add data to a batch for later writing to the Google Sheets document.

        :param sheet_name: The name of the sheet where data should be written.
        :param cell_range: The cell range where data should be written (e.g., 'A1').
        :param value: The value to write to the cell.
        :param function_name: The name of the function adding data to the batch.
        """
        if function_name not in self.batch_data:
            self.batch_data[function_name] = []
        self.batch_data[function_name].append(
            {"sheet_name": sheet_name, "cell_range": cell_range, "value": value}
        )

    def write_to_sheet(self, sheet_name, cell_range, value, function_name):
        """
        Write data to a specific cell in a sheet.

        :param sheet_name: The name of the sheet where data should be written.
        :param cell_range: The cell range where data should be written (e.g., 'A1').
        :param value: The value to write to the cell.
        :param function_name: The name of the function writing the data.
        """
        self.add_data_to_batch(sheet_name, cell_range, value, function_name)

    def write_batch(self):
        """
        Write the accumulated batch data to the Google Sheets document.
        """
        # Initialize batch values
        batch_values = []
        for function_name, data_list in self.batch_data.items():
            for data in data_list:
                value = [[data["value"]]]
                batch_values.append(
                    {
                        "range": f"{data['sheet_name']}!{data['cell_range']}",
                        "values": value,
                    }
                )

        # Clear the batch data after processing
        self.batch_data = {}

        body = {"data": batch_values, "valueInputOption": "RAW"}
        try:
            self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheet_id, body=body
            ).execute()
            print("Batch data written successfully.")
        except HttpError as e:
            print(f"Error writing batch to spreadsheet: {e}")


def get_and_write_team_names(fantasy_league, manager, sheet_name):
    """
    Get and write team names to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write team names to.
    """
    teams_data = fantasy_league.get_teams_data()
    index_row = 1

    for team in teams_data:
        cell_range = f"{chr(ord('A') + team.index_column)}{index_row}"
        manager.write_to_sheet(
            sheet_name,
            cell_range,
            team.team.team_name,
            "print_team_names_to_sheet",
        )
    cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
    manager.write_to_sheet(
        sheet_name,
        cell_range,
        "League average",
        "print_team_names_to_sheet",
    )
    manager.write_batch()


def get_and_write_average_points_per_player(
    fantasy_league,
    manager,
    sheet_name,
    start_position,
    positions_to_check,
    positions_to_omit,
    started_or_not,
):
    """
    Get and write average points per player to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write average points per player.
    :param start_position: The starting position in the sheet for writing.
    :param positions_to_check: List of positions to check (e.g., ["QB", "WR"]).
    :param positions_to_omit: List of positions to omit (e.g., ["BE"]).
    :param started_or_not: Indicator for player status (e.g., "started" or "overall").
    """
    teams_data = fantasy_league.average_points_per_player(
        positions_to_check, positions_to_omit
    )
    position_index = 0

    # Iterate over positions
    for index_row, position in enumerate(positions_to_check, start=start_position):
        manager.write_to_sheet(
            sheet_name,
            f"{chr(ord('A'))}{index_row}",
            f"Pts/{position} {started_or_not}",
            "average_points_per_player",
        )
        league_average = 0
        for team in teams_data:
            league_average += team.get_average_score_per_player()[position_index][1]
            cell_range = f"{chr(ord('A') + team.index_column)}{index_row}"
            manager.add_data_to_batch(
                sheet_name,
                cell_range,
                team.get_average_score_per_player()[position_index][1],
                "average_points_per_player",
            )
        cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            round(league_average / len(teams_data), 2),
            "average_points_per_player",
        )
        position_index += 1  # Move to the next position

    manager.write_batch()


def get_and_write_average_points_per_player_all_weeks(
    fantasy_league,
    weeks,
    manager,
    sheet_name,
    start_position,
    positions_to_check,
    positions_to_omit,
    started_or_not,
):
    """
    Get and write average points per player for all weeks to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param weeks: The number of weeks to consider.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write average points per player for all weeks.
    :param start_position: The starting position in the sheet for writing.
    :param positions_to_check: List of positions to check (e.g., ["QB", "WR"]).
    :param positions_to_omit: List of positions to omit (e.g., ["BE"]).
    :param started_or_not: Indicator for player status (e.g., "started" or "overall").
    """
    overall_teams_data = []
    position_index = 0

    for week in range(1, weeks + 1):
        teams_data = fantasy_league.average_points_per_player(
            positions_to_check, positions_to_omit, week
        )
        sorted_teams_data = sorted(
            teams_data, key=lambda team_data: team_data.index_column
        )

        if not overall_teams_data:
            overall_teams_data = [
                team.get_average_score_per_player() for team in sorted_teams_data
            ]
        else:
            for i, team in enumerate(sorted_teams_data):
                if len(overall_teams_data[i]) != len(
                    team.get_average_score_per_player()
                ):
                    # Handle varying tuple lengths
                    print(
                        f"Warning: Varying tuple lengths for team {team.team.team_name}"
                    )
                    continue
                overall_teams_data[i] = [
                    (pos, val1 + val2)
                    for (pos, val1), (_, val2) in zip(
                        overall_teams_data[i],
                        team.get_average_score_per_player(),
                    )
                ]
                average_overall_team_data = []
                for sub_list in overall_teams_data:
                    new_sub_list = [
                        (position, value / weeks, 2) for position, value in sub_list
                    ]
                    average_overall_team_data.append(new_sub_list)
    for index_row, position in enumerate(positions_to_check, start=start_position):
        manager.write_to_sheet(
            sheet_name,
            f"{chr(ord('A'))}{index_row}",
            f"Pts/{position} {started_or_not}",
            "average_points_per_player_all_weeks",
        )
        league_average = 0
        for team in range(len(teams_data)):
            cell_range = f"{chr(ord('A') + team + 1)}{index_row}"
            league_average += round(
                average_overall_team_data[team][position_index][1], 2
            )
            manager.add_data_to_batch(
                sheet_name,
                cell_range,
                round(average_overall_team_data[team][position_index][1], 2),
                "average_points_per_player_all_weeks",
            )
        cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            round(league_average / len(teams_data), 2),
            "average_points_per_player_all_weeks",
        )

        position_index += 1  # Move to the next position

    manager.write_batch()


def get_and_write_power_ranking(fantasy_league, manager, sheet_name, index_row):
    """
    Get and write power rankings to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write power rankings.
    :param index_row: The row index where power rankings should be written.
    """
    teams_data = fantasy_league.power_ranking_per_player()

    # Iterate over positions
    manager.write_to_sheet(
        sheet_name,
        f"{chr(ord('A'))}{index_row}",
        f"Power Ranking",
        "power_ranking",
    )
    league_average = 0
    for team in teams_data:
        cell_range = f"{chr(ord('A') + team.index_column)}{index_row}"
        league_average += float(team.get_power_ranking())
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            team.get_power_ranking(),
            "power_ranking",
        )
    cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
    manager.add_data_to_batch(
        sheet_name,
        cell_range,
        league_average / len(teams_data),
        "power_ranking",
    )

    manager.write_batch()


def get_and_write_power_ranking_all_weeks(
    fantasy_league, weeks, manager, sheet_name, index_row
):
    """
    Get and write power rankings for all weeks to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param weeks: The number of weeks to consider.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write power rankings for all weeks.
    :param index_row: The row index where power rankings should be written.
    """
    overall_teams_data = []
    for week in range(1, weeks + 1):
        teams_data = fantasy_league.power_ranking_per_player(week)
        sorted_teams_data = sorted(
            teams_data, key=lambda team_data: team_data.index_column
        )
        if not overall_teams_data:
            overall_teams_data = [
                float(team.get_power_ranking()) for team in sorted_teams_data
            ]
        else:
            for i, team_data in enumerate(sorted_teams_data):
                power_ranking = team_data.get_power_ranking()
                overall_teams_data[i] += float(power_ranking)
    rounded_values = [round(value / weeks, 2) for value in overall_teams_data]

    manager.write_to_sheet(
        sheet_name,
        f"{chr(ord('A'))}{index_row}",
        "Power Ranking",
        "power_ranking",
    )
    league_average = 0
    for position_index, value in enumerate(rounded_values, start=1):
        league_average += value
        cell_range = f"{chr(ord('A') + position_index)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            value,
            "power_ranking",
        )
    cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
    manager.add_data_to_batch(
        sheet_name,
        cell_range,
        round(league_average / len(teams_data), 2),
        "power_ranking",
    )

    manager.write_batch()


def get_and_write_over_under_projection(fantasy_league, manager, sheet_name, index_row):
    """
    Get and write over/under projections to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write over/under projections.
    :param index_row: The row index where over/under projections should be written.
    """
    teams_data = fantasy_league.get_teams_data()

    manager.write_to_sheet(
        sheet_name,
        f"{chr(ord('A'))}{index_row}",
        f"Over/Under Projection",
        "projection_diff",
    )
    league_average = 0
    for team in teams_data:
        score_diff = round(team.score - team.projected, 2)
        cell_range = f"{chr(ord('A') + team.index_column)}{index_row}"
        league_average += score_diff
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            score_diff,
            "power_ranking",
        )
    cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
    manager.add_data_to_batch(
        sheet_name,
        cell_range,
        round(league_average / len(teams_data), 2),
        "power_ranking",
    )

    manager.write_batch()


def get_and_write_over_under_projection_all_weeks(
    fantasy_league, weeks, manager, sheet_name, index_row
):
    """
    Get and write over/under projections for all weeks to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param weeks: The number of weeks to consider.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write over/under projections for all weeks.
    :param index_row: The row index where over/under projections should be written.
    """
    overall_teams_data = []
    for week in range(1, weeks + 1):
        teams_data = fantasy_league.power_ranking_per_player(week)
        sorted_teams_data = sorted(
            teams_data, key=lambda team_data: team_data.index_column
        )
        if not overall_teams_data:
            overall_teams_data = [
                round(team.score - team.projected, 2) for team in sorted_teams_data
            ]
        else:
            for i, team_data in enumerate(sorted_teams_data):
                overall_teams_data[i] += round(team_data.score - team_data.projected, 2)
    rounded_values = [round(value / weeks, 2) for value in overall_teams_data]

    manager.write_to_sheet(
        sheet_name,
        f"{chr(ord('A'))}{index_row}",
        f"Over/Under Projection",
        "projection_diff",
    )
    league_average = 0
    for position_index, value in enumerate(rounded_values, start=1):
        league_average += value
        cell_range = f"{chr(ord('A') + position_index)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            value,
            "power_ranking",
        )
    cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
    manager.add_data_to_batch(
        sheet_name,
        cell_range,
        round(league_average / len(teams_data), 2),
        "power_ranking",
    )

    manager.write_batch()


def get_and_write_percentage_of_points_per_position(
    fantasy_league,
    manager,
    sheet_name,
    start_position,
    positions_to_check,
    positions_to_omit,
    started_or_not,
):
    """
    Get and write percentage of points per position to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write percentage of points per position.
    :param start_position: The starting position in the sheet for writing.
    :param positions_to_check: List of positions to check (e.g., ["QB", "WR"]).
    :param positions_to_omit: List of positions to omit (e.g., ["BE"]).
    :param started_or_not: Indicator for player status (e.g., "Started").
    """
    teams_data = fantasy_league.percentage_of_points_by_position(
        positions_to_check, positions_to_omit
    )

    position_index = 0

    # Iterate over positions
    for index_row, position in enumerate(positions_to_check, start=start_position):
        manager.write_to_sheet(
            sheet_name,
            f"{chr(ord('A'))}{index_row}",
            f"% points/{position} {started_or_not}",
            "percentage_of_points_per_position",
        )
        league_average = 0
        for team in teams_data:
            league_average += round(
                team.get_percentage_per_position()[position_index][1], 2
            )
            cell_range = f"{chr(ord('A') + team.index_column)}{index_row}"
            manager.add_data_to_batch(
                sheet_name,
                cell_range,
                round(team.get_percentage_per_position()[position_index][1], 2),
                "percentage_of_points_per_position",
            )
        position_index += 1  # Move to the next position
        cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            round(league_average / len(teams_data), 2),
            "percentage_of_points_per_position",
        )

    manager.write_batch()


def get_and_write_percentage_of_points_per_position_all_weeks(
    fantasy_league,
    weeks,
    manager,
    sheet_name,
    start_position,
    positions_to_check,
    positions_to_omit,
    started_or_not,
):
    """
    Get and write percentage of points per position for all weeks to the specified Google Sheets document.

    :param fantasy_league: The fantasy football league data.
    :param weeks: The number of weeks to consider.
    :param manager: An instance of GoogleSheetsManager.
    :param sheet_name: The name of the sheet to write percentage of points per position for all weeks.
    :param start_position: The starting position in the sheet for writing.
    :param positions_to_check: List of positions to check (e.g., ["QB", "WR"]).
    :param positions_to_omit: List of positions to omit (e.g., ["BE"]).
    :param started_or_not: Indicator for player status (e.g., "Started").
    """
    overall_teams_data = []
    position_index = 0

    for week in range(1, weeks + 1):
        teams_data = fantasy_league.percentage_of_points_by_position(
            positions_to_check, positions_to_omit, week
        )
        sorted_teams_data = sorted(
            teams_data, key=lambda team_data: team_data.index_column
        )
        if not overall_teams_data:
            overall_teams_data = [
                team.get_percentage_per_position() for team in sorted_teams_data
            ]
        else:
            for i, team in enumerate(sorted_teams_data):
                if len(overall_teams_data[i]) != len(
                    team.get_percentage_per_position()
                ):
                    # Handle varying tuple lengths
                    print(
                        f"Warning: Varying tuple lengths for team {team.team.team_name}"
                    )
                    continue
                overall_teams_data[i] = [
                    (pos, val1 + val2)
                    for (pos, val1), (_, val2) in zip(
                        overall_teams_data[i],
                        team.get_percentage_per_position(),
                    )
                ]
                average_overall_team_data = []
                for sub_list in overall_teams_data:
                    new_sub_list = [
                        (position, value / weeks) for position, value in sub_list
                    ]
                    average_overall_team_data.append(new_sub_list)
    for index_row, position in enumerate(positions_to_check, start=start_position):
        manager.write_to_sheet(
            sheet_name,
            f"{chr(ord('A'))}{index_row}",
            f"% points/{position} {started_or_not}",
            "percentage_of_points_per_position",
        )
        league_average = 0
        for team in range(len(teams_data)):
            league_average += average_overall_team_data[team][position_index][1]
            cell_range = f"{chr(ord('A') + team + 1)}{index_row}"
            manager.add_data_to_batch(
                sheet_name,
                cell_range,
                round(average_overall_team_data[team][position_index][1], 2),
                "percentage_of_points_per_position",
            )
        position_index += 1  # Move to the next position
        cell_range = f"{chr(ord('A') + len(teams_data) + 1)}{index_row}"
        manager.add_data_to_batch(
            sheet_name,
            cell_range,
            round(league_average / len(teams_data), 2),
            "percentage_of_points_per_position",
        )

    manager.write_batch()


def run_analysis(week):
    credentials_file = "credentials.json"
    with open(credentials_file, "r") as json_file:
        credentials = json.load(json_file)
    spreadsheet_id = credentials["spreadsheet_id"]
    league_creator = LeagueCreator(
        league_id=credentials["league_id"],
        year=2023,
        espn_s2=credentials["espn_s2"],
        swid=credentials["swid"],
    )
    batch_size = 15
    overall_sheet_name = "Summary"
    # Create the League instance using the factory
    league = league_creator.create_league()

    manager = GoogleSheetsManager(credentials_file, spreadsheet_id, batch_size)
    fantasy_league = FantasyFootballLeague(league, week)
    print("we here")
    if not manager.check_sheet_exists(overall_sheet_name):
        print("YO")
        manager.create_new_sheet(overall_sheet_name)
        manager.delete_other_sheets(manager, overall_sheet_name)
    get_and_write_team_names(fantasy_league, manager, overall_sheet_name)
    get_and_write_average_points_per_player_all_weeks(
        fantasy_league,
        week,
        manager,
        overall_sheet_name,
        2,
        ["QB", "WR", "RB", "TE", "K", "D/ST"],
        ["BE"],
        "started",
    )

    get_and_write_average_points_per_player_all_weeks(
        fantasy_league,
        week,
        manager,
        overall_sheet_name,
        9,
        ["QB", "WR", "RB", "TE", "K", "D/ST"],
        [],
        "overall",
    )

    get_and_write_power_ranking_all_weeks(
        fantasy_league, week, manager, overall_sheet_name, 16
    )

    get_and_write_over_under_projection_all_weeks(
        fantasy_league, week, manager, overall_sheet_name, 18
    )

    get_and_write_percentage_of_points_per_position_all_weeks(
        fantasy_league,
        week,
        manager,
        overall_sheet_name,
        20,
        ["QB", "WR", "RB", "TE", "K", "D/ST"],
        ["BE"],
        "Started",
    )
    get_and_write_percentage_of_points_per_position_all_weeks(
        fantasy_league,
        week,
        manager,
        overall_sheet_name,
        27,
        ["QB", "WR", "RB", "TE", "K", "D/ST"],
        [],
        "Overall",
    )
    for week in range(1, week + 1):
        fantasy_league = FantasyFootballLeague(league, week)
        sheet_name = f"Week {week}"
        if not manager.check_sheet_exists(sheet_name):
            manager.create_new_sheet(sheet_name)

            get_and_write_team_names(fantasy_league, manager, sheet_name)

            get_and_write_average_points_per_player(
                fantasy_league,
                manager,
                sheet_name,
                2,
                ["QB", "WR", "RB", "TE", "K", "D/ST"],
                ["BE"],
                "started",
            )

            get_and_write_average_points_per_player(
                fantasy_league,
                manager,
                sheet_name,
                9,
                ["QB", "WR", "RB", "TE", "K", "D/ST"],
                [],
                "overall",
            )
            get_and_write_power_ranking(fantasy_league, manager, sheet_name, 16)

            get_and_write_over_under_projection(fantasy_league, manager, sheet_name, 18)

            get_and_write_percentage_of_points_per_position(
                fantasy_league,
                manager,
                sheet_name,
                20,
                ["QB", "WR", "RB", "TE", "K", "D/ST"],
                ["BE"],
                "Started",
            )
            get_and_write_percentage_of_points_per_position(
                fantasy_league,
                manager,
                sheet_name,
                27,
                ["QB", "WR", "RB", "TE", "K", "D/ST"],
                [],
                "Overall",
            )


def lambda_handler(event, context):
    first_week = "2023-09-5"

    if first_week:
        # Convert the "first_week" string to a datetime object
        first_week_date = datetime.datetime.strptime(first_week, "%Y-%m-%d")

        # Get the current date
        current_date = datetime.datetime.now()
        # Calculate the difference in weeks between first_week_date and current_date
        weeks_difference = (current_date - first_week_date).days // 7
        print(weeks_difference)
        run_analysis(weeks_difference)
