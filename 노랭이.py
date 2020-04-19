import discord
import openpyxl
import random

client = discord.Client()
admin_role = [651413680306651188, 642987212916391956, 651405902967996427, 642987402146873382, 661468035441360943]

@client.event
async def on_ready():
    print(client.user.id, "준비됐음다")
    print("https://discordapp.com/oauth2/authorize?client_id=655650874005127168permissions=8&scope=bot")

    game = discord.Game("지구 안전 수칙서 읽는 중...")
    await client.change_presence(status=discord.Status.online, activity=game)


@client.event
async def on_message(message):
    if message.content.startswith("노랭아 ") or message.content.startswith("ㄴ "):
        if message.content.startswith("노랭아 도움") or message.content.startswith("ㄴ 도움"):
            embed = discord.Embed(
                title="**<노랭이의 도움말>**",
                description="TOP님을 위해 chang06과 자연화맛 음료수가 만든 봇이닷\n\n",
                colour=discord.Colour.green()
            )
            embed.set_thumbnail(
                url="https://cdn.discordapp.com/attachments/651417391091417094/651726452714242059/a2e406895be038cb.png")
            embed.add_field(name='--------------------------__<관리자 전용 명령어>__-------------------------', value=
            "> 노랭아 공지 [채널 아이디] (내용)\n"
            + "노랭이가 채널 위치에 공지를 적는다.\n", inline=False)
            embed.add_field(name='-----------------------------__<공용 명령어>__-----------------------------', value=
             "> 노랭아 [아무말]\n"
            + "노랭이가 배운 말들 중 그에 맞는 대답을 한다\n"
            + "> 노랭아 기억해 [A]/[B]\n"
            + "누군가 [A]라고 말하면 스쿨봇이 [B]라고 대답한다\n"
            + "예시 : `노랭아 말추가 안녕/ㅎㅇㅎㅇ`\n", inline=False)

            await message.channel.send(embed=embed)







        elif message.content.startswith("노랭아 기억해") or message.content.startswith("ㄴ 기억해"):
            if str(message.author.id) == "492222241501741056" or str(message.author.id) == "481077311890784276" or str(
                    message.author.id) == "550226629717131287" or str(message.author.id) == "542503495752876035" or str(
                message.author.id) == "477280468400865280":
              file = openpyxl.load_workbook("노랭이.xlsx")
              sheet = file.active

              q = message.content.split("/")[0]
              a = message.content.split("/")[1]

              if message.content.startswith("노랭아 "):
                  q2 = q[8:]
              elif message.content.startswith("ㄴ "):
                  q2 = q[6:]

              i = 1
              flag = 0
              while sheet["A" + str(i)].value != None:
                  if sheet["A" + str(i)].value == q2: flag = 1
                  i += 1

              if flag == 0:
                  sheet["A" + str(i)].value = str(q2)
                  sheet["B" + str(i)].value = str(a)
                  sheet["C" + str(i)].value = str(message.author.id)
                  sheet["D" + str(i)].value = str(message.author)

                  await message.channel.send("[" + str(q2) + "]라고 하면 [" + str(a) + "]라고 대답하는 것을 배웠어!")
                  file.save("노랭이.xlsx")
              else:
                  embed = discord.Embed(
                      title="**<Error! 삐비빅>**",
                      description="가소로운 뇨오석~\n그건 이미 배운 말이라굿!\n강해저서 돌아와라 닝겐(킹시국<<ㅍㅍ)",
                      colour=discord.Colour.red()
                  )
                  embed.set_thumbnail(
                      url="https://cdn.discordapp.com/attachments/651417391091417094/651726452714242059/a2e406895be038cb.png")
                  await message.channel.send(embed=embed)
            else:
                embed = discord.Embed(
                    title="**<Error! 삐비빅>**",
                    description="가소로운 뇨오석~\n이 명령어는 관리자만 쓸 수 있다굿!\n강해저서 돌아와라 닝겐(킹시국<<ㅍㅍ)",
                    colour=discord.Colour.red()
                )
                embed.set_thumbnail(
                    url="https://cdn.discordapp.com/attachments/651417391091417094/651726452714242059/a2e406895be038cb.png")
                await message.channel.send(embed=embed)



        elif message.content.startswith("ㄴ 공지"):
            if str(message.author.id) == "492222241501741056" or str(message.author.id) == "481077311890784276" or str(
                    message.author.id) == "550226629717131287" or str(message.author.id) == "542503495752876035" or str(message.author.id) == "477280468400865280":

                channel = message.content[5:23]
                msg = message.content[24:]
                await client.get_channel(int(channel)).send(msg)
            else:
                embed = discord.Embed(
                    title="**<Error! 삐비빅>**",
                    description="가소로운 뇨오석~\n이 명령어는 관리자만 쓸 수 있다굿!\n강해저서 돌아와라 닝겐(킹시국<<ㅍㅍ)",
                    colour=discord.Colour.red()
                )
                embed.set_thumbnail(
                    url="https://cdn.discordapp.com/attachments/651417391091417094/651726452714242059/a2e406895be038cb.png")
                await message.channel.send(embed=embed)





        elif message.content.startswith("노랭아 공지"):
            if str(message.author.id) == "492222241501741056" or str(message.author.id) == "481077311890784276" or str(
                    message.author.id) == "550226629717131287" or str(message.author.id) == "542503495752876035" or str(message.author.id) == "477280468400865280":

                channel = message.content[5:23]
                msg = message.content[24:]
                await client.get_channel(int(channel)).send(msg)
            else:
                embed = discord.Embed(
                    title="**<Error! 삐비빅>**",
                    description="가소로운 뇨오석~\n이 명령어는 관리자만 쓸 수 있다굿!\n강해저서 돌아와라 닝겐(킹시국<<ㅍㅍ)",
                    colour=discord.Colour.red()
                )
                embed.set_thumbnail(
                    url="https://cdn.discordapp.com/attachments/651417391091417094/651726452714242059/a2e406895be038cb.png")
                await message.channel.send(embed=embed)

        elif message.content.startswith("노랭아 치킨은?") or message.content.startswith("ㄴ 치킨은?"):
            await message.channel.send("닭둘기")

        elif message.content.startswith("노랭아 닭둘기는?") or message.content.startswith("ㄴ 닭둘기는?"):

            await message.channel.send("닭둘기")

        elif message.content.startswith("노랭아 너는?") or message.content.startswith("ㄴ 너는?"):
            await message.channel.send("달에서 온 노랭이랍니다")

        elif message.content.startswith("노랭아 먹어도 돼?") or message.content.startswith("ㄴ 먹어도 돼?"):
            await message.channel.send("뭐를 먹는 다하는 지는 모르겠지만,나를 먹으면 너를 들고서 바닷가로 뛰어가서 던져 버릴꺼얌")

        elif message.content.startswith("노랭아 달은?") or message.content.startswith("ㄴ 달은?"):
            await message.channel.send("GO향")

        elif message.content.startswith("노랭아 토끼는?") or message.content.startswith("ㄴ 토끼는?"):
           await message.channel.send("만무늬 토기? ㅏㅏ 번지수 잘못 찾음 미안")

        elif message.content.startswith("노랭아 나는?") or message.content.startswith("ㄴ 나는?"):
            await message.channel.send("안 물어 봤는 뎁")

        elif message.content.startswith("노랭아 죽어") or message.content.startswith("ㄴ 죽어"):
            await message.channel.send("그 나이를 먹고서도 그런 말 입에 담고 사십니까...? 불쌍하네요")

        elif message.content.startswith("노랭아 뒤져") or message.content.startswith("ㄴ 뒤져"):
            await message.channel.send("앞★정")

        elif message.content.startswith("노랭아 배고파") or message.content.startswith("ㄴ 배고파"):
            await message.channel.send("음..배고픈건 알겠는데 난 먹으면 안대..")

        elif message.content.startswith("노랭아 불은?") or message.content.startswith("ㄴ 불은?"):
            await message.channel.send("살려주세요 형님")

        elif message.content.startswith("노랭아 우주선은?") or message.content.startswith("ㄴ 우주선은?"):
            await message.channel.send("내가 온도조절장치 터트린 기계")

        elif message.content.startswith("노랭아 이거 알아?") or message.content.startswith("ㄴ 이거 알아?"):
            await message.channel.send("어 몰라")

        elif message.content.startswith("노랭아 곰은?") or message.content.startswith("ㄴ 곰은?"):
            await message.channel.send("문")

        elif message.content.startswith("노랭아 가위바위보") or message.content.startswith("ㄴ 가위바위보"):
            await message.channel.send("탕")

        elif message.content.startswith("노랭아 노랭파이") or message.content.startswith("ㄴ 노랭파이"):
            pic = message.content.split(" ")[4]
            await message.channel.send(file=discord.File(pic))

        elif message.content.startswith("노랭아 가위바위보") or message.content.startswith("ㄴ 가위바위보"):
            await message.channel.send("탕")

























        else:
            file = openpyxl.load_workbook("노랭이.xlsx")
            sheet = file.active

            if message.content.startswith("노랭아 "):
                q = message.content[4:]
            elif message.content.startswith("ㄴ "):
                q = message.content[2:]

            i = 1
            flag = 0
            while sheet["A" + str(i)].value != None:
                if sheet["A" + str(i)].value == q:
                    await message.channel.send(str(sheet["B" + str(i)].value))
                    flag = 1

                i += 1

            if flag == 0: await message.channel.send("님이 뭐라고 하시는지 잘 모루게쒀요\n`노랭아 도움`을 입력하셔서 명령어를 확인해주세욥")


client.run("NzAwNTAwMjA1ODU4MTkzNDU5.XpldNQ.6WpD7KPOHIG_e2sDeQCE5gUwWQ0")