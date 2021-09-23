import loginPage from '../../Page objects/Login.page';
import itemsPlanningPlanningPage, {
  PlanningCreateUpdate,
  PlanningRowObject,
} from '../../Page objects/ItemsPlanning/ItemsPlanningPlanningPage';
import itemsPlanningModalPage from '../../Page objects/ItemsPlanning/ItemsPlanningModal.page';
import {
  generateRandmString,
  getRandomInt,
} from '../../Helpers/helper-functions';
import myEformsPage from '../../Page objects/MyEforms.page';
import foldersPage from '../../Page objects/Folders.page';
import { format, parse } from 'date-fns';

const expect = require('chai').expect;
let planningData: PlanningCreateUpdate = {
  name: [generateRandmString(), generateRandmString(), generateRandmString()],
  eFormName: generateRandmString(),
  description: generateRandmString(),
  repeatEvery: '1',
  repeatType: 'Dag',
  repeatUntil: new Date(2020, 6, 10),
  folderName: generateRandmString(),
  type: generateRandmString(),
  buildYear: '10',
  locationCode: '12345',
  startFrom: new Date(2020, 7, 9),
  number: '10',
  pushMessageEnabled: false,
  daysBeforeRedeploymentPushMessage: getRandomInt(1, 27),
};
let folderNameForEdit = generateRandmString();
let eFormNameForEdit = generateRandmString();
describe('Items planning actions - Edit', function () {
  before(async () => {
    await loginPage.open('/auth');
    await loginPage.login();
    if (await myEformsPage.rowNum() >= 2) {
      planningData.eFormName = (await myEformsPage.getEformRowObj(1)).eFormName;
      eFormNameForEdit = (await myEformsPage.getEformRowObj(2)).eFormName;
    } else {
      // Create eforms
      if (await myEformsPage.rowNum() === 1) {
        planningData.eFormName = (await myEformsPage.getEformRowObj(1)).eFormName;
      } else {
        await myEformsPage.createNewEform(planningData.eFormName);
      }
      await myEformsPage.createNewEform(eFormNameForEdit);
    }

    await myEformsPage.Navbar.goToFolderPage();
    if (await foldersPage.rowNum() >= 2) {
      planningData.folderName = (await foldersPage.getFolder(1)).name;
      folderNameForEdit = (await foldersPage.getFolder(2)).name;
    } else {
      // Create two folder
      if (await foldersPage.rowNum() === 1) {
        planningData.folderName = (await foldersPage.getFolder(1)).name;
      } else {
        await foldersPage.createNewFolder(planningData.folderName, 'Description');
      }
      await foldersPage.createNewFolder(folderNameForEdit, 'Description');
    }
    await itemsPlanningPlanningPage.goToPlanningsPage();
  });
  it('should create a new planning', async () => {
    await itemsPlanningModalPage.createPlanning(planningData);
  });
  it('should change all fields after edit', async () => {
    let planningRowObject = await itemsPlanningPlanningPage.getPlaningByName(
      planningData.name[0]
    );
    const tempForSwapFolderName = planningData.folderName;
    const tempForSwapEFormFormName = planningData.eFormName;
    planningData = {
      name: [
        generateRandmString(),
        generateRandmString(),
        generateRandmString(),
      ],
      repeatType: 'Uge',
      description: generateRandmString(),
      folderName: folderNameForEdit,
      eFormName: eFormNameForEdit,
      number: '2',
      startFrom: parse('7/3/2020', 'M/d/yyyy', new Date()),
      locationCode: '54321',
      buildYear: '20',
      type: generateRandmString(),
      repeatUntil: parse('10/18/2020', 'M/d/yyyy', new Date()),
      repeatEvery: '2',
      pushMessageEnabled: true,
      daysBeforeRedeploymentPushMessage: getRandomInt(1, 27),
    };
    folderNameForEdit = tempForSwapFolderName;
    eFormNameForEdit = tempForSwapEFormFormName;
    await planningRowObject.update(planningData);

    // Check that list is edited successfully in table
    planningRowObject = await itemsPlanningPlanningPage.getPlaningByName(
      planningData.name[0]
    );
    await planningRowObject.openEdit();
    browser.pause(1000);
    for (let i = 0; i < planningData.name.length; i++) {
      expect(
        await (await itemsPlanningModalPage.editPlanningItemName(i)).getValue(),
        'Name save is incorrect'
      ).eq(planningData.name[i]);
    }
    expect(
      await (await itemsPlanningModalPage.editPlanningDescription()).getValue(),
      'Description save is incorrect'
    ).eq(planningData.description);
    expect(
      await (await (await itemsPlanningModalPage.editPlanningSelector()).$('.ng-value')).getText(),
      'Saved template is incorrect'
    ).eq(planningData.eFormName);
    expect(
      await (await itemsPlanningModalPage.editRepeatEvery()).getValue(),
      'Saved repeat every is incorrect'
    ).eq(planningData.repeatEvery);
    expect(
      format(
        parse(
          await (await itemsPlanningModalPage.editRepeatUntil()).getValue(),
          'M/d/yyyy',
          new Date()
        ),
        'M/d/yyyy'
      ),
      'Saved repeat until is incorrect'
    ).eq(format(planningData.repeatUntil, 'M/d/yyyy'));
    expect(
      await (await (await itemsPlanningModalPage.editRepeatType()).$('.ng-value-label')).getText(),
      'Saved repeat type is incorrect'
    ).eq(planningData.repeatType);
    expect(
      await (await itemsPlanningModalPage.editItemType()).getValue(),
      'Saved type is incorrect'
    ).eq(planningData.type);
    expect(
      await (await itemsPlanningModalPage.editItemBuildYear()).getValue(),
      'Saved build year is incorrect'
    ).eq(planningData.buildYear);
    expect(
      await (await (await itemsPlanningModalPage.editFolderName())
        .$('#editFolderSelectorInput'))
        .getValue(),
      'Saved folder name is incorrect'
    ).eq(planningData.folderName);
    expect(
      await (await itemsPlanningModalPage.editItemLocationCode()).getValue(),
      'Saved location code is incorrect'
    ).eq(planningData.locationCode);
    expect(
      format(
        parse(
          await (await itemsPlanningModalPage.editStartFrom()).getValue(),
          'M/d/yyyy',
          new Date()
        ),
        'M/d/yyyy'
      ),
      'Saved start from is incorrect'
    ).eq(format(planningData.startFrom, 'M/d/yyyy'));
    expect(
      await (await (await itemsPlanningModalPage.pushMessageEnabledEdit())
        .$('.ng-value-label'))
        .getText(),
      'Saved pushMessageEnabled code is incorrect'
    ).eq(
      planningData.pushMessageEnabled
        ? 'Aktiveret'
        : 'Deaktiveret'
    );
    expect(
      +await (await (await itemsPlanningModalPage.editDaysBeforeRedeploymentPushMessage())
        .$('.ng-value-label'))
        .getText(),
      'Saved editDaysBeforeRedeploymentPushMessage code is incorrect'
    ).eq(planningData.daysBeforeRedeploymentPushMessage);
    await PlanningRowObject.closeEdit(true);
  });
  after(async () => {
    // Delete
    const planningRowObject = await itemsPlanningPlanningPage.getPlaningByName(
      planningData.name[0]
    );
    await planningRowObject.delete();

    await myEformsPage.Navbar.goToFolderPage();
    await (await foldersPage.getFolderByName(planningData.folderName)).delete();
    await (await foldersPage.getFolderByName(folderNameForEdit)).delete();

    await myEformsPage.Navbar.goToMyEForms();
    await (await myEformsPage.getEformsRowObjByNameEForm(eFormNameForEdit)).deleteEForm();
    await (await myEformsPage
      .getEformsRowObjByNameEForm(planningData.eFormName))
      .deleteEForm();
  });
});
