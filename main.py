import discord
from discord.ext import commands, tasks
from discord import app_commands
import json
import os
from keep_alive import keep_alive
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
from typing import Dict, List, Optional, Tuple
import asyncio
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from collections import defaultdict

load_dotenv()

# ==================== 設定區 ====================
TOKEN = os.getenv("DISCORD_TOKEN")

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Asia/Taipei")
except Exception:
    TZ = timezone(timedelta(hours=8))

JOBS = ["鐵衣", "血河", "九靈", "素問", "神相", "龍吟", "碎夢", "玄機", "潮光"]

JOB_EMOJI = {
    "鐵衣": "🟠",
    "血河": "🔴",
    "素問": "🩷",
    "碎夢": "🔵",
    "龍吟": "🟢",
    "玄機": "🟡",
    "九靈": "🟣",
    "神相": "🔷",
    "潮光": "💙",
}

MAX_PLAYERS = 60
MAX_PER_JOB = 14
NUM_TEAMS = 10
SLOTS_PER_TEAM = 6

TEAM_NAMES = [
    "一團一隊", "一團保/拆", "一團保鑣",
    "二團一隊", "二團保/拆", "二團保鑣",
    "三團一隊", "三團二隊", "三團三隊", "三團四隊"
]

# 請改成你伺服器實際的身分組名稱
REQUIRED_ROLE_NAME = "優樂城【荷官】"

DATA_FILE = "event_data.json"
WEEKDAY_ZH = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
# ================================================

intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.guilds = True

bot = commands.Bot(command_prefix="!", intents=intents)


def load_data() -> dict:
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data: dict):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_guild_data(guild_id: int) -> dict:
    data = load_data()
    gid = str(guild_id)
    if gid not in data:
        data[gid] = {
            "event_name": None,
            "event_date": None,
            "opponent": None,
            "deadline": None,
            "deadline_iso": None,
            "channel_id": None,
            "thread_id": None,
            "message_id": None,
            "signups": {},
            "teams": [[None] * SLOTS_PER_TEAM for _ in range(NUM_TEAMS)],
            "registration_closed": False
        }
        save_data(data)
    if "teams" not in data[gid]:
        data[gid]["teams"] = [[None] * SLOTS_PER_TEAM for _ in range(NUM_TEAMS)]
    return data[gid]


def update_guild_data(guild_id: int, guild_data: dict):
    data = load_data()
    data[str(guild_id)] = guild_data
    save_data(data)


def format_datetime_display(dt: datetime) -> str:
    weekday = WEEKDAY_ZH[dt.weekday()]
    return f"{dt.strftime('%Y-%m-%d')}【{weekday}】 {dt.strftime('%H:%M')}"


def parse_yyyymmdd(date_str: str, hour: int, minute: int) -> Optional[datetime]:
    try:
        date_str = date_str.strip().replace("-", "")
        if len(date_str) != 8:
            return None
        year = int(date_str[0:4])
        month = int(date_str[4:6])
        day = int(date_str[6:8])
        return datetime(year, month, day, hour, minute, tzinfo=TZ)
    except Exception:
        return None


def is_deadline_passed(guild_data: dict) -> bool:
    if guild_data.get("registration_closed"):
        return True
    deadline_iso = guild_data.get("deadline_iso")
    if not deadline_iso:
        return False
    try:
        deadline = datetime.fromisoformat(deadline_iso)
        return datetime.now(TZ) >= deadline
    except Exception:
        return False


def get_job_counts(signups: dict) -> Dict[str, int]:
    counts = {job: 0 for job in JOBS}
    for info in signups.values():
        job = info.get("job")
        if job in counts:
            counts[job] += 1
    return counts


def get_remaining_slots(signups: dict) -> int:
    return MAX_PLAYERS - len(signups)


def count_user_signups(signups: dict, user_id: str) -> Tuple[int, int]:
    normal = 0
    proxy = 0
    for uid, info in signups.items():
        owner = info.get("owner_id", uid)
        if owner == user_id:
            if info.get("is_proxy"):
                proxy += 1
            else:
                normal += 1
    return normal, proxy


def find_best_slot(guild_data: dict, job: str) -> Optional[Tuple[int, int]]:
    teams = guild_data["teams"]

    def team_job_count(t_idx: int, j: str) -> int:
        return sum(1 for p in teams[t_idx] if p and p.get("job") == j)

    def empty_slots(t_idx: int) -> List[int]:
        return [i for i, p in enumerate(teams[t_idx]) if p is None]

    candidates = []

    for t in range(NUM_TEAMS):
        empties = empty_slots(t)
        if not empties:
            continue

        # 職業不重複（素問最多2個）
        if job != "素問" and team_job_count(t, job) >= 1:
            continue
        if job == "素問" and team_job_count(t, "素問") >= 2:
            continue

        # 位置偏好
        if job in ("鐵衣", "血河"):
            preferred = [0] if 0 in empties else empties
        elif job == "素問":
            preferred = [s for s in (4, 5) if s in empties] or empties
        else:
            preferred = empties

        if not preferred:
            continue

        slot = preferred[0]
        score = len(empties) * 10

        # 強制需求加權
        if t < 4:  # 1-4 隊：鐵衣 + 2素問
            if job == "鐵衣" and team_job_count(t, "鐵衣") == 0:
                score += 50
            if job == "素問" and team_job_count(t, "素問") < 2:
                score += 40
        elif t in (4, 5):  # 5-6 隊：至少1素問
            if job == "素問" and team_job_count(t, "素問") == 0:
                score += 45
        else:  # 7-10 隊：血河 + 素問
            if job == "血河" and team_job_count(t, "血河") == 0:
                score += 50
            if job == "素問" and team_job_count(t, "素問") == 0:
                score += 40

        candidates.append((score, t, slot))

    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][2]


def place_player(guild_data: dict, key: str, info: dict) -> bool:
    job = info["job"]
    result = find_best_slot(guild_data, job)
    if result is None:
        return False
    t_idx, s_idx = result
    info["team"] = t_idx
    info["slot"] = s_idx
    guild_data["teams"][t_idx][s_idx] = {
        "uid": key,
        "char_name": info["char_name"],
        "job": job,
        "is_proxy": info.get("is_proxy", False),
        "name": info.get("name", "")
    }
    return True


def remove_player_from_teams(guild_data: dict, key: str):
    for t in range(NUM_TEAMS):
        for s in range(SLOTS_PER_TEAM):
            p = guild_data["teams"][t][s]
            if p and p.get("uid") == key:
                guild_data["teams"][t][s] = None


def format_status_embed(guild_data: dict) -> discord.Embed:
    signups = guild_data.get("signups", {})
    job_counts = get_job_counts(signups)
    remaining = get_remaining_slots(signups)
    total = len(signups)
    closed = is_deadline_passed(guild_data)

    color = discord.Color.red() if closed or remaining <= 0 else discord.Color.blue()
    embed = discord.Embed(
        title=f"⚔️ {guild_data.get('event_name', '聯賽活動')} 報名狀態",
        color=color
    )
    embed.add_field(name="🏰 對手幫會", value=guild_data.get("opponent") or "未設定", inline=True)
    embed.add_field(name="⏰ 約戰時間", value=guild_data.get("event_date") or "未設定", inline=True)
    embed.add_field(name="📅 報名截止", value=guild_data.get("deadline") or "未設定", inline=True)

    status_text = "🔴 **報名已截止**" if closed else f"🟢 報名進行中（剩餘 **{remaining}** 名額）"
    embed.add_field(name="📊 目前人數", value=f"{total} / {MAX_PLAYERS}\n{status_text}", inline=False)

    job_lines = []
    for job in JOBS:
        count = job_counts[job]
        emoji = JOB_EMOJI.get(job, "⚪")
        status = "🔴 已滿" if count >= MAX_PER_JOB else f"剩餘 {MAX_PER_JOB - count}"
        job_lines.append(f"{emoji} **{job}**：{count}/{MAX_PER_JOB}　{status}")
    embed.add_field(name="🗡️ 職業報名狀況", value="\n".join(job_lines), inline=False)

    if signups:
        by_job = {job: [] for job in JOBS}
        for info in signups.values():
            job = info.get("job")
            char_name = info.get("char_name") or "未知"
            prefix = "(代報)" if info.get("is_proxy") else ""
            if job in by_job:
                by_job[job].append(f"{prefix}{char_name}【{job}】")
        list_lines = []
        for job in JOBS:
            if by_job[job]:
                emoji = JOB_EMOJI.get(job, "")
                list_lines.append(f"{emoji} **{job}**\n" + "、".join(by_job[job]))
        full_list = "\n\n".join(list_lines)
        if len(full_list) > 1000:
            full_list = full_list[:980] + "\n…（已截斷）"
        embed.add_field(name="👥 已報名成功人員", value=full_list, inline=False)
    else:
        embed.add_field(name="👥 已報名成功人員", value="目前尚無人報名", inline=False)

    embed.set_footer(text="⛔ 報名已截止" if closed else "請使用下方按鈕進行報名 / 代報")
    embed.timestamp = datetime.now(TZ)
    return embed


def format_team_status_embed(guild_data: dict) -> discord.Embed:
    teams = guild_data.get("teams", [[None]*6 for _ in range(10)])
    job_counts = get_job_counts(guild_data.get("signups", {}))

    embed = discord.Embed(
        title=f"📋 {guild_data.get('event_name', '聯賽')} 隊伍分配狀態",
        color=discord.Color.gold()
    )

    for t_idx in range(NUM_TEAMS):
        team = teams[t_idx]
        lines = []
        for s_idx, p in enumerate(team):
            pos = s_idx + 1
            if p is None:
                lines.append(f"`{pos}.` （空）")
            else:
                emoji = JOB_EMOJI.get(p["job"], "⚪")
                prefix = "(代報)" if p.get("is_proxy") else ""
                lines.append(f"`{pos}.` {emoji} {prefix}{p['char_name']}【{p['job']}】")
        embed.add_field(name=f"🛡️ {TEAM_NAMES[t_idx]}", value="\n".join(lines), inline=True)

    stats = [f"{JOB_EMOJI.get(j, '⚪')} {j}:{job_counts[j]}" for j in JOBS]
    embed.add_field(name="📊 職業統計", value="　".join(stats), inline=False)
    embed.set_footer(text="使用 /move_member 可調整成員位置")
    embed.timestamp = datetime.now(TZ)
    return embed


async def update_thread_status(guild: discord.Guild):
    guild_data = get_guild_data(guild.id)
    thread_id = guild_data.get("thread_id")
    message_id = guild_data.get("message_id")
    if not thread_id or not message_id:
        return
    try:
        thread = guild.get_thread(thread_id) or await bot.fetch_channel(thread_id)
    except Exception:
        return

    closed = is_deadline_passed(guild_data)
    if closed and not guild_data.get("registration_closed"):
        guild_data["registration_closed"] = True
        update_guild_data(guild.id, guild_data)

    embed = format_status_embed(guild_data)
    view = SignupView(disabled=closed)
    try:
        msg = await thread.fetch_message(message_id)
        await msg.edit(embed=embed, view=view)
    except Exception:
        new_msg = await thread.send(embed=embed, view=view)
        guild_data["message_id"] = new_msg.id
        update_guild_data(guild.id, guild_data)


class CharacterNameModal(discord.ui.Modal, title="輸入角色名稱"):
    def __init__(self, job: str, is_proxy: bool):
        super().__init__()
        self.job = job
        self.is_proxy = is_proxy
        self.char_name_input = discord.ui.TextInput(
            label="角色名稱" + ("（代報）" if is_proxy else ""),
            placeholder="請輸入遊戲內的角色名稱",
            min_length=1,
            max_length=20,
            required=True
        )
        self.add_item(self.char_name_input)

    async def on_submit(self, interaction: discord.Interaction):
        await complete_signup(interaction, self.job, self.char_name_input.value.strip(), self.is_proxy)


async def complete_signup(interaction: discord.Interaction, job: str, char_name: str, is_proxy: bool):
    guild = interaction.guild
    if not guild:
        await interaction.response.send_message("❌ 只能在伺服器內使用。", ephemeral=True)
        return

    guild_data = get_guild_data(guild.id)
    user_id = str(interaction.user.id)

    if is_deadline_passed(guild_data):
        await interaction.response.send_message("⛔ 報名已截止。", ephemeral=True)
        return

    normal_cnt, proxy_cnt = count_user_signups(guild_data["signups"], user_id)
    if is_proxy:
        if proxy_cnt >= 1:
            await interaction.response.send_message("❌ 你已經代報過一次了。", ephemeral=True)
            return
    else:
        if normal_cnt >= 1:
            await interaction.response.send_message("❌ 你已經報名過一次了。", ephemeral=True)
            return

    if len(guild_data["signups"]) >= MAX_PLAYERS:
        await interaction.response.send_message("❌ 已達總人數上限 60 人。", ephemeral=True)
        return

    job_counts = get_job_counts(guild_data["signups"])
    if job_counts.get(job, 0) >= MAX_PER_JOB:
        await interaction.response.send_message(f"❌ **{job}** 已達上限 14 人。", ephemeral=True)
        return

    info = {
        "job": job,
        "char_name": char_name,
        "name": str(interaction.user),
        "is_proxy": is_proxy,
        "owner_id": user_id,
        "joined_at": datetime.now(TZ).isoformat()
    }

    key = user_id if not is_proxy else f"proxy_{user_id}_{len(guild_data['signups'])}"
    success = place_player(guild_data, key, info)
    if not success:
        info["team"] = None
        info["slot"] = None

    guild_data["signups"][key] = info
    update_guild_data(guild.id, guild_data)
    await update_thread_status(guild)

    remaining = get_remaining_slots(guild_data["signups"])
    team_text = ""
    if info.get("team") is not None:
        team_text = f"\n隊伍：{TEAM_NAMES[info['team']]} 第 {info['slot']+1} 位"

    prefix = "(代報)" if is_proxy else ""
    await interaction.response.send_message(
        f"✅ **{'代報' if is_proxy else '報名'}成功！**\n"
        f"角色：{prefix}**{char_name}**\n職業：**{job}**{team_text}\n剩餘名額：{remaining}",
        ephemeral=True
    )

    try:
        thread_id = guild_data.get("thread_id")
        thread_link = f"https://discord.com/channels/{guild.id}/{thread_id}" if thread_id else ""
        dm = discord.Embed(
            title="🎉 報名成功通知" if not is_proxy else "📝 代報成功通知",
            description=f"{'你已成功報名' if not is_proxy else '你已成功代報'} **{guild_data['event_name']}**",
            color=discord.Color.green()
        )
        dm.add_field(name="角色名稱", value=f"{prefix}{char_name}", inline=True)
        dm.add_field(name="職業", value=job, inline=True)
        if team_text:
            dm.add_field(name="分配隊伍", value=team_text.strip(), inline=False)
        if thread_link:
            dm.add_field(name="📢 討論串", value=f"[點此前往]({thread_link})", inline=False)
        await interaction.user.send(embed=dm)
    except discord.Forbidden:
        pass

    await asyncio.sleep(7)
    try:
        await interaction.delete_original_response()
    except Exception:
        pass


async def handle_cancel(interaction: discord.Interaction):
    guild = interaction.guild
    if not guild:
        return
    guild_data = get_guild_data(guild.id)
    user_id = str(interaction.user.id)

    if is_deadline_passed(guild_data):
        await interaction.response.send_message("⛔ 報名已截止，無法取消。", ephemeral=True)
        return

    to_remove = [key for key, info in guild_data["signups"].items()
                 if key == user_id or info.get("owner_id") == user_id]

    if not to_remove:
        await interaction.response.send_message("❌ 你目前沒有任何報名紀錄。", ephemeral=True)
        return

    removed_names = []
    for key in to_remove:
        info = guild_data["signups"].pop(key)
        remove_player_from_teams(guild_data, key)
        removed_names.append(f"{'(代報)' if info.get('is_proxy') else ''}{info.get('char_name')}")

    update_guild_data(guild.id, guild_data)
    await update_thread_status(guild)

    await interaction.response.send_message(
        f"✅ 已取消以下報名：\n" + "\n".join(removed_names), ephemeral=True
    )

    try:
        dm = discord.Embed(title="❌ 取消報名通知", color=discord.Color.orange(),
                           description=f"你已取消 **{guild_data.get('event_name')}** 的報名")
        await interaction.user.send(embed=dm)
    except discord.Forbidden:
        pass

    await asyncio.sleep(6)
    try:
        await interaction.delete_original_response()
    except Exception:
        pass


class JobSelectView(discord.ui.View):
    def __init__(self, is_proxy: bool):
        super().__init__(timeout=60)
        self.is_proxy = is_proxy
        options = [discord.SelectOption(label=job, value=job) for job in JOBS]
        select = discord.ui.Select(placeholder="請選擇職業", options=options)
        select.callback = self.on_select
        self.add_item(select)

    async def on_select(self, interaction: discord.Interaction):
        job = interaction.data["values"][0]
        await interaction.response.send_modal(CharacterNameModal(job, self.is_proxy))


class SignupView(discord.ui.View):
    def __init__(self, disabled: bool = False):
        super().__init__(timeout=None)

        btn_signup = discord.ui.Button(label="報名", style=discord.ButtonStyle.primary,
                                       custom_id="btn_signup", emoji="📝", disabled=disabled)
        btn_signup.callback = self.signup_callback
        self.add_item(btn_signup)

        btn_proxy = discord.ui.Button(label="代報", style=discord.ButtonStyle.secondary,
                                      custom_id="btn_proxy", emoji="👥", disabled=disabled)
        btn_proxy.callback = self.proxy_callback
        self.add_item(btn_proxy)

        btn_cancel = discord.ui.Button(label="取消報名", style=discord.ButtonStyle.danger,
                                       custom_id="btn_cancel", emoji="⚠️", disabled=disabled)
        btn_cancel.callback = self.cancel_callback
        self.add_item(btn_cancel)

        btn_refresh = discord.ui.Button(label="重新整理", style=discord.ButtonStyle.secondary,
                                        custom_id="btn_refresh", emoji="🔄")
        btn_refresh.callback = self.refresh_callback
        self.add_item(btn_refresh)

    async def signup_callback(self, interaction: discord.Interaction):
        if is_deadline_passed(get_guild_data(interaction.guild.id)):
            await interaction.response.send_message("⛔ 報名已截止。", ephemeral=True)
            return
        await interaction.response.send_message("請選擇要報名的職業：",
                                                view=JobSelectView(is_proxy=False), ephemeral=True)

    async def proxy_callback(self, interaction: discord.Interaction):
        if is_deadline_passed(get_guild_data(interaction.guild.id)):
            await interaction.response.send_message("⛔ 報名已截止。", ephemeral=True)
            return
        await interaction.response.send_message("請選擇要**代報**的職業：",
                                                view=JobSelectView(is_proxy=True), ephemeral=True)

    async def cancel_callback(self, interaction: discord.Interaction):
        await handle_cancel(interaction)

    async def refresh_callback(self, interaction: discord.Interaction):
        await update_thread_status(interaction.guild)
        await interaction.response.send_message("✅ 已重新整理", ephemeral=True)
        await asyncio.sleep(3)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass


def has_required_role():
    async def predicate(interaction: discord.Interaction) -> bool:
        if interaction.user.guild_permissions.administrator:
            return True
        role = discord.utils.get(interaction.guild.roles, name=REQUIRED_ROLE_NAME)
        if role and role in interaction.user.roles:
            return True
        await interaction.response.send_message(
            f"❌ 你需要擁有 **{REQUIRED_ROLE_NAME}** 身分組才能使用此指令。", ephemeral=True)
        return False
    return app_commands.check(predicate)


HOUR_CHOICES = [app_commands.Choice(name=f"{h:02d} 時", value=h) for h in range(24)]
MINUTE_CHOICES = [
    app_commands.Choice(name="00 分（整點）", value=0),
    app_commands.Choice(name="30 分", value=30),
]


@bot.tree.command(name="new_event", description="【管理員】建立新的聯賽活動")
@app_commands.describe(
    名稱="活動名稱",
    對手幫會="對手幫會名稱",
    約戰日期="約戰日期（格式 YYYYMMDD，例如 20260920）",
    約戰時="約戰小時",
    約戰分="約戰分鐘",
    截止日期="報名截止日期（格式 YYYYMMDD）",
    截止時="截止小時",
    截止分="截止分鐘"
)
@app_commands.choices(約戰時=HOUR_CHOICES, 約戰分=MINUTE_CHOICES, 截止時=HOUR_CHOICES, 截止分=MINUTE_CHOICES)
@app_commands.default_permissions(administrator=True)
@has_required_role()
async def new_event(
    interaction: discord.Interaction,
    名稱: str, 對手幫會: str,
    約戰日期: str, 約戰時: app_commands.Choice[int], 約戰分: app_commands.Choice[int],
    截止日期: str, 截止時: app_commands.Choice[int], 截止分: app_commands.Choice[int]
):
    await interaction.response.defer(ephemeral=True)

    battle_dt = parse_yyyymmdd(約戰日期, 約戰時.value, 約戰分.value)
    if not battle_dt:
        await interaction.followup.send("❌ 約戰日期格式錯誤，請使用 YYYYMMDD（例如 20260920）", ephemeral=True)
        return
    deadline_dt = parse_yyyymmdd(截止日期, 截止時.value, 截止分.value)
    if not deadline_dt:
        await interaction.followup.send("❌ 截止日期格式錯誤，請使用 YYYYMMDD", ephemeral=True)
        return
    if deadline_dt >= battle_dt:
        await interaction.followup.send("❌ 報名截止時間必須早於約戰時間", ephemeral=True)
        return

    battle_display = format_datetime_display(battle_dt)
    deadline_display = format_datetime_display(deadline_dt)

    guild = interaction.guild
    channel = interaction.channel
    if isinstance(channel, discord.Thread):
        channel = channel.parent
    if not isinstance(channel, discord.TextChannel):
        await interaction.followup.send("❌ 請在文字頻道使用", ephemeral=True)
        return

    guild_data = get_guild_data(guild.id)
    guild_data.update({
        "event_name": 名稱,
        "opponent": 對手幫會,
        "event_date": battle_display,
        "deadline": deadline_display,
        "deadline_iso": deadline_dt.isoformat(),
        "signups": {},
        "teams": [[None] * SLOTS_PER_TEAM for _ in range(NUM_TEAMS)],
        "registration_closed": False,
        "channel_id": channel.id
    })

    try:
        thread = await channel.create_thread(
            name=f"【{名稱}】vs {對手幫會}｜{battle_display}",
            type=discord.ChannelType.public_thread,
            auto_archive_duration=10080
        )
        guild_data["thread_id"] = thread.id

        # 禁止一般成員傳訊息（只能點按鈕）
        overwrite = thread.overwrites_for(guild.default_role)
        overwrite.send_messages = False
        await thread.set_permissions(guild.default_role, overwrite=overwrite)

        embed = format_status_embed(guild_data)
        view = SignupView(disabled=False)
        msg = await thread.send(
            content=(
                f"📢 **{名稱} 報名討論串**\n"
                f"🏰 對手：{對手幫會}\n"
                f"⏰ 約戰：{battle_display}\n"
                f"📅 截止：{deadline_display}\n\n"
                f"請使用下方按鈕「報名」或「代報」\n"
                f"（本討論串已禁止一般訊息，僅能透過按鈕操作）"
            ),
            embed=embed, view=view
        )
        guild_data["message_id"] = msg.id
        update_guild_data(guild.id, guild_data)
        await interaction.followup.send(f"✅ 活動已建立！\n討論串：{thread.mention}", ephemeral=True)
    except Exception as e:
        await interaction.followup.send(f"❌ 建立失敗：{e}", ephemeral=True)


@bot.tree.command(name="export_excel", description="【管理員】匯出報名名單 Excel（含分組）")
@app_commands.default_permissions(administrator=True)
@has_required_role()
async def export_excel(interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)
    guild_data = get_guild_data(interaction.guild.id)
    signups = guild_data.get("signups", {})
    teams = guild_data.get("teams", [])

    if not signups:
        await interaction.followup.send("❌ 目前沒有報名資料", ephemeral=True)
        return

    wb = Workbook()
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws = wb.active
    ws.title = "分組名單"
    ws.merge_cells('A1:E1')
    ws['A1'] = f"{guild_data.get('event_name')} vs {guild_data.get('opponent')}｜{guild_data.get('event_date')}"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["隊伍", "位置", "角色名稱", "職業", "代報"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin

    row = 4
    for t_idx, team in enumerate(teams):
        for s_idx, p in enumerate(team):
            if p is None:
                values = [TEAM_NAMES[t_idx], s_idx+1, "（空）", "", ""]
            else:
                values = [TEAM_NAMES[t_idx], s_idx+1, p.get("char_name", ""), p.get("job", ""),
                          "是" if p.get("is_proxy") else "否"]
            for c, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=c, value=v)
                cell.border = thin
                cell.alignment = Alignment(horizontal="center")
            row += 1
        row += 1

    for col in range(1, 6):
        ws.column_dimensions[get_column_letter(col)].width = 16

    ws2 = wb.create_sheet("完整名單")
    for col, h in enumerate(["序號", "角色名稱", "職業", "代報", "Discord名稱", "ID"], 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    for idx, (key, info) in enumerate(signups.items(), 1):
        ws2.append([idx, info.get("char_name"), info.get("job"),
                    "是" if info.get("is_proxy") else "否", info.get("name"), key])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    filename = f"{guild_data.get('event_name', 'event')}_名單_{datetime.now(TZ).strftime('%Y%m%d_%H%M')}.xlsx"
    await interaction.followup.send(
        content=f"📊 已匯出（共 {len(signups)} 人）",
        file=discord.File(fp=buffer, filename=filename), ephemeral=True
    )


@bot.tree.command(name="event_status", description="查看目前隊伍分配狀態")
@has_required_role()
async def event_status(interaction: discord.Interaction):
    guild_data = get_guild_data(interaction.guild.id)
    if not guild_data.get("event_name"):
        await interaction.response.send_message("❌ 目前沒有進行中的活動", ephemeral=True)
        return
    embed = format_team_status_embed(guild_data)
    await interaction.response.send_message(embed=embed, ephemeral=True)


@bot.tree.command(name="close_event", description="【管理員】強制關閉報名")
@app_commands.default_permissions(administrator=True)
@has_required_role()
async def close_event(interaction: discord.Interaction):
    guild_data = get_guild_data(interaction.guild.id)
    guild_data["registration_closed"] = True
    update_guild_data(interaction.guild.id, guild_data)
    await update_thread_status(interaction.guild)
    await interaction.response.send_message("✅ 已強制關閉報名", ephemeral=True)


@bot.tree.command(name="move_member", description="【管理員】調整成員到指定隊伍與位置")
@app_commands.describe(
    角色名稱="要移動的角色名稱",
    目標隊伍="1~10（對應一團一隊～三團四隊）",
    目標位置="1~6"
)
@app_commands.default_permissions(administrator=True)
@has_required_role()
async def move_member(
    interaction: discord.Interaction,
    角色名稱: str,
    目標隊伍: app_commands.Range[int, 1, 10],
    目標位置: app_commands.Range[int, 1, 6]
):
    await interaction.response.defer(ephemeral=True)
    guild_data = get_guild_data(interaction.guild.id)
    teams = guild_data["teams"]
    signups = guild_data["signups"]

    found_key = None
    found_info = None
    for key, info in signups.items():
        if info.get("char_name") == 角色名稱:
            found_key = key
            found_info = info
            break

    if not found_info:
        await interaction.followup.send(f"❌ 找不到角色「{角色名稱}」", ephemeral=True)
        return

    t_idx = 目標隊伍 - 1
    s_idx = 目標位置 - 1

    # 清空目標位置原本的人
    old_p = teams[t_idx][s_idx]
    if old_p:
        for k, inf in signups.items():
            if inf.get("char_name") == old_p.get("char_name"):
                inf["team"] = None
                inf["slot"] = None

    remove_player_from_teams(guild_data, found_key)

    teams[t_idx][s_idx] = {
        "uid": found_key,
        "char_name": found_info["char_name"],
        "job": found_info["job"],
        "is_proxy": found_info.get("is_proxy", False),
        "name": found_info.get("name", "")
    }
    found_info["team"] = t_idx
    found_info["slot"] = s_idx

    update_guild_data(interaction.guild.id, guild_data)
    await update_thread_status(interaction.guild)

    await interaction.followup.send(
        f"✅ 已將 **{角色名稱}** 移動到 **{TEAM_NAMES[t_idx]}** 的第 **{目標位置}** 位置",
        ephemeral=True
    )


@tasks.loop(minutes=1)
async def check_deadline_task():
    data = load_data()
    for gid, gdata in data.items():
        if gdata.get("registration_closed"):
            continue
        if is_deadline_passed(gdata):
            gdata["registration_closed"] = True
            update_guild_data(int(gid), gdata)
            guild = bot.get_guild(int(gid))
            if guild:
                await update_thread_status(guild)


@bot.event
async def on_ready():
    print(f"✅ 已登入為 {bot.user}")
    bot.add_view(SignupView(disabled=False))
    try:
        synced = await bot.tree.sync()
        print(f"✅ 同步 {len(synced)} 個指令")
    except Exception as e:
        print(e)
    if not check_deadline_task.is_running():
        check_deadline_task.start()


if __name__ == "__main__":
    try:
        from keep_alive import keep_alive
        keep_alive()
        print("✅ keep_alive 已啟動")
    except ImportError:
        print("⚠️ 找不到 keep_alive.py，略過")

    token = os.getenv("DISCORD_TOKEN") or ""
    if not token:
        raise Exception("Please set DISCORD_TOKEN")

    try:
        bot.run(token)
    except discord.HTTPException as e:
        if e.status == 429:
            print("The Discord servers denied the connection for making too many requests")
            print("Get help from https://stackoverflow.com/questions/66724687/in-discord-py-how-to-solve-the-error-for-toomanyrequests")
        else:
            raise e
