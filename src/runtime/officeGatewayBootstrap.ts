import type { OfficeGateway } from "../platform/office/officeGateway";
import type { CommandId, CommandResult } from "../shared/types";

type OfficeGatewayBootstrapOptions<Gateway extends Pick<OfficeGateway, "runWrite">> = {
  createGateway(): Gateway;
  ready(): Promise<unknown>;
};

export const createOfficeGatewayProvider = <
  Gateway extends Pick<OfficeGateway, "runWrite">,
>({
  createGateway,
  ready,
}: OfficeGatewayBootstrapOptions<Gateway>): (() => Promise<Gateway>) => {
  let gateway: Promise<Gateway> | null = null;
  return async () => {
    const pending = gateway ?? ready().then(() => createGateway());
    gateway = pending;
    try {
      return await pending;
    } catch (error) {
      if (gateway === pending) gateway = null;
      throw error;
    }
  };
};

export const createOfficeCommandRunner = ({
  createGateway,
  ready,
}: OfficeGatewayBootstrapOptions<Pick<OfficeGateway, "runWrite">>): ((id: CommandId) => Promise<CommandResult>) => {
  const getGateway = createOfficeGatewayProvider({ createGateway, ready });
  return async (id) => (await getGateway()).runWrite({ id });
};
